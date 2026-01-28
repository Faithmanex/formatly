# Formatly V6 🎓

Formatly V6 is a comprehensive, AI-powered academic document formatting ecosystem. It leverages advanced Large Language Models (Gemini, HuggingFace) to intelligently parse, structure, and format Word documents according to strict academic standards (APA, MLA, Chicago, Harvard).

This project consists of three main interfaces powered by a shared core library:
1.  **Command Line Interface (CLI)**: For quick, scriptable formatting tasks.
2.  **Desktop Application**: A modern, user-friendly GUI built with PySide6.
3.  **REST API**: A scalable FastAPI backend for integration and remote processing.

## ✨ Key Features

-   **🤖 AI-Powered Structure Detection**: Uses LLMs (Gemini, HuggingFace) to intelligently identify document elements (Abstract, Headings, References, etc.) without relying on rigid templates.
-   **📚 Multi-Style Support**: Full support for APA, MLA, Chicago, and Harvard citation styles.
-   **📝 Intelligent Formatting**: Automatically applies fonts, margins, spacing, indentation, and heading levels.
-   **🔄 Track Changes**: Optional mode to generate documents with all formatting changes tracked for review.
-   **✅ Advanced Proofreading**: Integrated AI spelling, grammar, and style checking.
-   **⚡ Dynamic & Resilient**: Features rate limit management, retry logic, and dynamic chunking for handling large documents.
-   **🖥️ Cross-Platform**: Works on Windows, macOS, and Linux.

---

## 🏗️ Architecture

The project is organized into a modular structure:

-   **`core/`**: Contains the business logic.
    -   `formatter.py`: Main formatting engine.
    -   `api_clients.py`: Clients for Gemini, HuggingFace, etc.
-   **`api/`**: FastAPI implementation (`api.py`) handling document queues, storage (Supabase), and processing.
-   **`desktop/`**: PySide6 desktop application (`GUI_pyside.py`) wrapping the core logic in a convenient UI.
-   **`app.py`**: The CLI entry point.
-   **`style_guides.py`**: Definitions for citation styles (margins, fonts, paragraph rules).

---

## 🚀 Installation

### Prerequisites

-   Python 3.8 or higher.
-   A Google Gemini API Key (or HuggingFace API Key depending on configuration).
-   (Optional) Supabase credentials if running the API.

### Setup Steps

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd Formatly-V6
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Environment Configuration:**
    Create a `.env` file in the root directory. You can copy `.env.example` if available.

    **Minimal `.env`:**
    ```env
    GEMINI_API_KEY=your_gemini_api_key_here
    GEMINI_MODEL=gemini-2.0-flash-exp
    ```

    **For API/Full Features:**
    ```env
    GEMINI_API_KEY=your_gemini_api_key
    HF_API_KEY=your_huggingface_key
    SUPABASE_URL=your_supabase_url
    SUPABASE_SERVICE_ROLE_KEY=your_service_key
    SUPABASE_JWT_SECRET=your_jwt_secret
    SUPABASE_ANON_KEY=your_anon_key
    ```

---

## 📖 Usage

### 1. Command Line Interface (CLI)

The CLI is perfect for batch processing or quick fixes.

**Basic Formatting:**
```bash
python app.py input_document.docx --style apa
```

**Interactive Mode:**
```bash
python app.py --interactive
```

**Auto-Fix Errors & Format:**
```bash
python app.py document.docx --fix-errors --style mla
```

**Generate Report Only:**
```bash
python app.py document.docx --report-only
```

**Available Options:**
-   `-s, --style`: `apa`, `mla`, `chicago`, `harvard` (default: `apa`)
-   `-t, --track-changes`: Enable tracked changes in the output.
-   `--english`: `us`, `gb`, `au`, `ca` (default: `us`).

### 2. Desktop Application

The desktop app provides a visual interface for selecting files, styles, and viewing progress.

**Launch:**
```bash
python desktop/GUI_pyside.py
```

*Note: The desktop app requires the `desktop/interface/` directory for its HTML/JS frontend assets.*

### 3. REST API

The API allows you to submit documents for processing programmatically. It integrates with Supabase for storage and authentication.

**Start the Server:**
```bash
python api.py
```
The API will run on `http://0.0.0.0:50000` (default).
Swagger documentation is available at `http://localhost:50000/docs`.

**Key Endpoints:**
-   `POST /api/documents/create-upload`: Generate an upload URL.
-   `POST /api/documents/upload-complete`: Trigger processing after upload.
-   `GET /api/documents/status/{job_id}`: Check processing status.

---

## 🎨 Supported Styles

| Style | Key Features |
| :--- | :--- |
| **APA** | Title Page, Abstract, Headings (Levels 1-5), Reference List, Times New Roman 12pt. |
| **MLA** | No Title Page (default), Heading on first page, Works Cited, Author-Page citation. |
| **Chicago** | Title Page, Bibliography, Footnotes/Endnotes support. |
| **Harvard** | Title Page, Reference List, Author-Year citation. |

---

## 📂 Project Structure

```
Formatly-V6/
├── api/                # API-related files
├── core/               # Core logic (formatting, AI clients)
│   ├── formatter.py
│   └── api_clients.py
├── desktop/            # Desktop GUI application
│   ├── GUI_pyside.py
│   └── interface/      # UI assets
├── utils/              # Utility scripts (spell check, rate limiting)
├── api.py              # FastAPI server entry point
├── app.py              # CLI entry point
├── config.py           # Configuration loader
├── style_guides.py     # Style definitions
├── requirements.txt    # Dependencies
└── README.md           # Documentation
```

## 🛠️ Development

-   **Testing**: Ensure you have valid API keys for testing AI components.
-   **Logs**: `formatting_stats.csv` tracks token usage and processing time.
-   **Formatting Logic**: Edit `style_guides.py` to adjust margins or font settings for specific styles.

## 📄 License

This project is licensed under the MIT License.
