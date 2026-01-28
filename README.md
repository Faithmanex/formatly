# Formatly 🎓

**Formatly** is an advanced, AI-powered academic document formatter designed to automate the process of applying citation styles (APA, MLA, Chicago, Harvard) to Word documents (`.docx`). It utilizes Google's Gemini AI (or Hugging Face models) to intelligently analyze document structure and apply rigorous formatting rules.

Formatly is available in three forms:
1.  **CLI (Command Line Interface)**: For quick, scriptable formatting tasks.
2.  **Desktop App**: A user-friendly GUI application.
3.  **API**: A FastAPI-based backend for integrating formatting services.

---

## ✨ Features

-   **🤖 AI-Powered Analysis**: Uses Large Language Models (LLMs) to intelligently detect document components (Headings, Abstracts, Lists, References) without relying on rigid templates.
-   **📚 Multi-Style Support**: Full support for **APA (7th Ed)**, **MLA (9th Ed)**, **Chicago**, and **Harvard** citation styles.
-   **🎯 Smart Formatting**: Automatically handles margins, fonts, line spacing, indentation, and page numbers according to the selected style guide.
-   **📝 Reference Management**: Sorts references, applies hanging indents, and formats citations.
-   **🔤 Spelling & Grammar**: Integrated AI-based spelling and grammar correction (CLI feature).
-   **🖥️ Cross-Platform**: Works on Windows, macOS, and Linux (Desktop app optimized for Windows).

---

## 🏗️ Architecture

The project is organized into modular components:

*   **`core/`**: The heart of Formatly. Contains the `AdvancedFormatter` logic, style definitions, and AI client wrappers. This logic is shared by the CLI and API.
*   **`api/`**: A FastAPI server that wraps the core logic, allowing for file uploads and processing via HTTP endpoints.
*   **`desktop/`**: A PySide6 (Qt) graphical interface. *Note: The desktop app currently maintains a self-contained copy of the core logic to ensure portability.*
*   **`app.py`**: The entry point for the Command Line Interface.

---

## 🚀 Installation

### Prerequisites

*   **Python 3.8+**
*   **Google Gemini API Key** (Get one [here](https://makersuite.google.com/app/apikey)) or a Hugging Face Token.

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/formatly.git
cd formatly
```

### 2. Install Dependencies

It is recommended to use a virtual environment.

```bash
# Create virtual environment
python -m venv venv

# Activate (Windows)
venv\Scripts\activate

# Activate (Mac/Linux)
source venv/bin/activate

# Install requirements
pip install -r requirements.txt
```

### 3. Configuration (.env)

Create a `.env` file in the root directory to store your API keys and configuration.

```env
# Required for AI Processing
GEMINI_API_KEY=your_gemini_api_key_here
GEMINI_MODEL=gemini-2.0-flash

# Optional: Hugging Face Fallback
HF_API_KEY=your_hugging_face_token

# API Server Config (Optional)
SUPABASE_URL=your_supabase_url
SUPABASE_SERVICE_ROLE_KEY=your_supabase_key
```

---

## 📖 Usage

### 1. Command Line Interface (CLI)

The CLI is the direct way to format documents.

**Basic Usage:**
```bash
python app.py input_document.docx
```

**Specify Output & Style:**
```bash
python app.py essay.docx --style mla --output formatted_essay.docx
```

**Interactive Mode:**
```bash
python app.py --interactive
```

**Available Options:**
*   `--style`: `apa` (default), `mla`, `chicago`, `harvard`
*   `--english`: `us` (default), `gb`, `ca`, `au` (for spell checking)
*   `--track-changes`: Enable change tracking in the output document.
*   `--fix-errors`: Auto-correct spelling and grammar before formatting.
*   `--report-only`: Generate a report of issues without modifying the file.

### 2. Desktop Application

To launch the Graphical User Interface:

```bash
python desktop/GUI_pyside.py
```
*   Select your document, choose a style, and click "Format".
*   The app uses a modern WebView interface powered by PySide6.

### 3. API Server

To start the local API server:

```bash
python api.py
```
*   The server will start (default port: 50000).
*   API Documentation (Swagger UI) available at: `http://localhost:50000/docs`

---

## 🎨 Supported Styles

| Style | Features |
| :--- | :--- |
| **APA 7** | Title Page, Abstract, Times New Roman 12pt, Double Spacing, Reference Hanging Indents. |
| **MLA 9** | No Title Page (Heading on Page 1), Works Cited, Double Spacing, Last Name + Page Header. |
| **Chicago** | Title Page, Footnotes/Endnotes support, Bibliography formatting. |
| **Harvard** | Title Page, Author-Date citations, Specific bibliography formatting. |

---

## 📁 Project Structure

```
Formatly/
├── app.py                  # CLI Entry Point
├── api.py                  # API Entry Point
├── style_guides.py         # Citation Style Definitions (Shared)
├── formatting_stats.csv    # Usage logs
├── core/                   # Shared Core Library
│   ├── formatter.py        # Main formatting engine
│   └── api_clients.py      # AI Model integrations
├── desktop/                # Desktop Application
│   ├── GUI_pyside.py       # GUI Entry Point
│   ├── interface/          # HTML/JS Front-end resources
│   └── core/               # Desktop-specific Core (Self-contained)
├── api/                    # Additional API resources
└── requirements.txt        # Dependencies
```

## 🛠️ Troubleshooting

*   **API Key Error**: Ensure `GEMINI_API_KEY` is set in your `.env` file.
*   **Permission Denied**: Ensure the input Word document is closed before running Formatly.
*   **"File not found"**: Provide the absolute path or place the file in the `documents/` folder.

---

**Made with ❤️ for academic excellence.**
