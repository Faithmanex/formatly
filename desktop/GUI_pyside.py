"""
Formatly Desktop Application (PySide6)
--------------------------------------
This module is the entry point for the desktop GUI application.
It uses QWebEngineView to render a modern HTML/JS interface while handling
backend logic in Python.

Architecture Note:
    - This application currently relies on a local copy of the core library
      located in `desktop/core/` and `desktop/style_guides.py`.
    - It does NOT currently import from the root `core/` directory to ensure
      self-contained distribution (e.g., with PyInstaller).
    - Future refactoring may aim to unify these codebases.
"""

import sys
import os
import json
import threading
import time
import winsound
from pathlib import Path
from dotenv import load_dotenv

from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog, QVBoxLayout, QWidget, QFrame
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtCore import QObject, Slot, Signal, QUrl, Qt, QPoint, QSize
from PySide6.QtGui import QIcon

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

class Bridge(QObject):
    """
    Bridge class exposed to JavaScript via QWebChannel.
    """
    # Define signals if needed (e.g. for progress updates)
    progressUpdated = Signal(str, int)
    formattingFinished = Signal(bool, str, str)
    notificationsUpdated = Signal() # Helper signal to trigger JS update
    aiResponseReceived = Signal(bool, str) # success, message
    aiResponseStart = Signal() # Signal to start a new system message
    aiResponseChunk = Signal(str) # Signal to stream content

    def __init__(self, window):
        super().__init__()
        self._window = window
        self._is_cancelled = False
        self._data = self._load_data()
        self._active_processing_count = 0
        
    @Slot(str)
    def ask_ai(self, user_message):
        """Handle chat messages from the UI using Groq."""
        def _chunk_handler(chunk):
            # Emit chunk signal
            try:
                self.aiResponseChunk.emit(chunk)
            except Exception as e:
                print(f"Error emitting chunk: {e}")

        def _run_ai_query():
            try:
                # 1. Initialize Groq Client
                from core.api_clients import GroqClient
                client = GroqClient()
                
                # 2. Define System Prompt
                system_prompt = (
                    "You are Formatly AI, an expert assistant for academic formatting, citation styles, and document structure. Provide helpful, accurate information about APA, MLA, Chicago, and other academic formats. Use markdown formatting for better readability, including code blocks for examples, bullet points for lists, and proper headings for organization."
                )
                
                # 3. Notify Start
                self.aiResponseStart.emit()

                # 4. Call API with streaming
                # capture response purely for logs/finalizing
                response_text, _ = client.generate_chat_response(
                    system_prompt, 
                    user_message, 
                    callback=_chunk_handler
                )
                
                # 5. Emit Final Success (Can be used to finalize the message or unlock input)
                # We pass the full text just in case, but UI might ignore it if it built it up via chunks
                self.aiResponseReceived.emit(True, response_text)
                
            except Exception as e:
                print(f"Chat Error: {e}")
                self.aiResponseReceived.emit(False, self._sanitize_error(e))

        threading.Thread(target=_run_ai_query, daemon=True).start()


    @Slot()
    def minimize_window(self):
        self._window.showMinimized()

    @Slot()
    def maximize_window(self):
        if self._window.isMaximized():
            self._window.restore()
        else:
            self._window.maximize()

    @Slot()
    def restore_window(self):
        self._window.restore()
    
    @Slot()
    def close_window(self):
        self._window.close()

    @Slot(result=str)
    def browse_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self._window, 
            "Select Document", 
            "", 
            "Word Documents (*.docx)"
        )
        return file_name if file_name else None

    @Slot(dict, result=dict)
    def start_formatting(self, options):
        self._is_cancelled = False
        self._active_processing_count += 1
        thread = threading.Thread(target=self._run_formatting, args=(options,))
        thread.daemon = True
        thread.start()
        return {"success": True}

    @Slot(result=dict)
    def cancel_formatting(self):
        self._is_cancelled = True
        return {"success": True}

    @Slot(str, result=dict)
    def open_file(self, path):
        try:
            if os.name == 'nt' and os.path.exists(path):
                os.startfile(path)
                return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    @Slot(result=list)
    def get_history(self):
        return self._data.get("history", [])

    @Slot(result=list)
    def get_notifications(self):
        return self._data.get("notifications", [])

    @Slot(int, result=dict)
    def mark_notification_read(self, notification_id):
        notifications = self._data.get("notifications", [])
        for n in notifications:
            if n.get("id") == notification_id:
                n["read"] = True
        self._save_data()
        return {"success": True}

    @Slot(result=dict)
    def clear_notifications(self):
        self._data["notifications"] = []
        self._save_data()
        return {"success": True}

    @Slot(result=dict)
    def get_preferences(self):
        return self._data.get("settings", {})

    @Slot(dict, result=dict)
    def save_preferences(self, prefs):
        self._data["settings"] = prefs
        self._save_data()
        return {"success": True}

    @Slot(result=dict)
    def get_stats(self):
        history = self._data.get("history", [])
        total_formatted = sum(1 for item in history if item.get("status") == "Completed")
        unique_documents = len(set(item.get("filename") for item in history))
        plan_usage_count = len(history) 
        
        return {
            "documents": unique_documents,
            "processing": self._active_processing_count,
            "formatted": total_formatted,
            "plan_usage": f"{plan_usage_count}/100"
        }

    @Slot(result=dict)
    def get_user_info(self):
        return {
            "name": "Faithman Founder",
            "plan": "Pro",
            "avatar": None
        }

    def _load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r") as f:
                    return json.load(f)
            except:
                pass
        return {"history": [], "settings": {}, "notifications": []}

    def _save_data(self):
        try:
            with open(DATA_FILE, "w") as f:
                json.dump(self._data, f, indent=2)
        except:
            pass

    def _add_to_history(self, filename, status, path, style="apa", size_bytes=0):
        try:
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
            self._data["history"] = self._data["history"][:500]
            self._save_data()
        except Exception as e:
            print(f"Error adding to history: {e}")

    def _add_notification(self, title, message, type="info"):
        notification = {
            "id": int(time.time() * 1000),
            "title": title,
            "message": message,
            "type": type,
            "date": time.strftime("%H:%M"),
            "read": False
        }
        self._data.setdefault("notifications", []).insert(0, notification)
        self._data["notifications"] = self._data["notifications"][:20]
        self._save_data()
        
        # Emit signal to notify JS
        # We need to do this carefully from a thread
        # In PySide signals are thread-safe
        self.notificationsUpdated.emit()

    def _run_formatting(self, options):
        try:
            # Need to handle path string safely
            input_path = Path(options['input'])
            
            if not input_path.exists():
                self.formattingFinished.emit(False, "Selected file does not exist", "")
                return
            
            try:
                file_size = input_path.stat().st_size
            except:
                file_size = 0
                
            style_name = validate_style(options['style'])
            output_path = determine_output_path(input_path, None, style_name)
            backend = options['backend']

            self.progressUpdated.emit("Initializing backend...", 10)

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
                self.formattingFinished.emit(False, "cancelled", "")
                return

            self.progressUpdated.emit("Analyzing and formatting document...", 40)
            english_variant = options.get('english_variant', 'us')
            formatter = AdvancedFormatter(style_name, ai_client=client, english_variant=english_variant)

            if style_name == 'mla':
                setattr(formatter, 'mla_heading', options.get('mla_heading', True))
                setattr(formatter, 'mla_title_page', options.get('mla_title_page', False))

            final_output_path_str = formatter.format_document(str(input_path), str(output_path))
            if final_output_path_str:
                output_path = Path(final_output_path_str)

            if self._is_cancelled:
                self.formattingFinished.emit(False, "cancelled", "")
                return

            if options.get('track_changes'):
                self.progressUpdated.emit("Tracking changes...", 80)
                tracker = TrackChanges(str(input_path.absolute()), str(output_path.absolute()))
                tracker.compare_docs()

            self.progressUpdated.emit("Success!", 100)
            
            self._add_to_history(input_path.name, "Completed", output_path, style=style_name, size_bytes=file_size)
            
            self._add_notification(
                "Formatting Complete", 
                f"{input_path.name} was successfully formatted in {style_name.upper()} style.", 
                "success"
            )
            
            winsound.MessageBeep(winsound.MB_OK)
            self.formattingFinished.emit(True, "Document formatted successfully!", str(output_path))

        except Exception as e:
            print(f"Formatting Error: {e}") # Log raw error to console
            friendly_msg = self._sanitize_error(e)
            self.formattingFinished.emit(False, friendly_msg, "")
            self._add_notification("Formatting Failed", friendly_msg, "error")
        finally:
            self._active_processing_count = max(0, self._active_processing_count - 1)

    def _sanitize_error(self, e):
        """Converts raw exceptions into user-friendly messages."""
        err_str = str(e).lower()
        
        if "timeout" in err_str:
            return "The operation timed out. The AI server is currently busy. Please try again in a few minutes."
        
        if "503" in err_str or "socket is null" in err_str or "connection" in err_str or "failed to connect" in err_str:
            return "Network Error: Unable to connect to the AI server. Please check your internet connection and firewall settings."
            
        if "api key" in err_str or "unauthorized" in err_str:
            return "Authorization Error: Invalid or missing API Key. Please check your settings."
            
        if "quota" in err_str or "limit" in err_str:
            return "Usage Limit Exceeded: You have reached the API rate limit. Please wait a moment before trying again."
            
        # Fallback for generic errors
        return f"An error occurred: {str(e)[:100]}{'...' if len(str(e)) > 100 else ''}"

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Formatly")
        self.resize(1200, 800)
        self.setMinimumSize(1000, 700)
        
        # Standard Window Setup
        # self.setWindowFlags(Qt.FramelessWindowHint) # Disabled by user request
        self.setStyleSheet("background-color: #050505;")
        
        # Central Widget & Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Add a border/shadow frame so it looks nice
        self.layout = QVBoxLayout(central_widget)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)
        
        # WebView
        self.browser = QWebEngineView()
        self.browser.setContextMenuPolicy(Qt.NoContextMenu) # Disable right click
        self.layout.addWidget(self.browser)

        # Setup Bridge
        self.channel = QWebChannel()
        self.bridge = Bridge(self)
        self.channel.registerObject("bridge", self.bridge)
        self.browser.page().setWebChannel(self.channel)

        # Load Content
        current_dir = os.path.dirname(os.path.abspath(__file__))
        html_path = os.path.join(current_dir, 'interface', 'index.html')
        self.browser.load(QUrl.fromLocalFile(html_path))

        # Window dragging logic
        self.old_pos = None

    # Window dragging logic removed as we are using native frame


    # Enable resizing on edges (simplified manual implementation)
    # For a robust solution, one would override native events (winEvent) handling WM_NCHITTEST
    # But for now let's hope standard resize works or we just use a small margin hack.
    # Actually, with FramelessWindowHint, standard resize borders are gone. 
    # We'll just rely on the user maximizing or using the fixed size for now or assume
    # they can use the Windows + Arrow keys. 
    # Implementing full manual resize in Python for Windows is verbose.

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()