# Formatly Development Guide 🛠️

This document provides an overview of the Formatly project architecture for developers. It explains how the different components (CLI, API, Desktop) are structured and interact with the core formatting logic.

## 🏗️ System Architecture

Formatly is built around a central "Core" library that handles the heavy lifting of document analysis (via AI) and formatting (via `python-docx`).

### High-Level Component View

```mermaid
graph TD
    User[User Input]

    subgraph "Core Logic (Shared)"
        Formatter[AdvancedFormatter]
        AI[AI Client (Gemini/HF)]
        Styles[Style Guides]
    end

    subgraph "Interfaces"
        CLI[CLI (app.py)]
        API[API Server (api.py)]
        Desktop[Desktop App (PySide6)]
    end

    User --> CLI
    User --> API
    User --> Desktop

    CLI --> Formatter
    API --> Formatter
    Desktop --> DesktopCore[Desktop Core (Local Copy)]
```

## 🧩 Component Details

### 1. Shared Core (`core/`)
The root `core/` directory contains the source of truth for the formatting engine.
-   **`formatter.py`**: Contains `AdvancedFormatter` and `DocumentStructureManager`. This is where the AI prompt generation and `python-docx` manipulation happen.
-   **`style_guides.py`**: A dictionary defining the rules (fonts, margins, spacing) for APA, MLA, Chicago, etc.
-   **`api_clients.py`**: Wrappers for Google Gemini and Hugging Face APIs.

### 2. Command Line Interface (`app.py`)
-   **Entry Point**: `app.py`
-   **Dependencies**: Imports directly from the root `core/` and `style_guides.py`.
-   **Functionality**: Handles argument parsing, local file I/O, and invokes the formatter.

### 3. API Server (`api.py`)
-   **Entry Point**: `api.py`
-   **Framework**: FastAPI
-   **Dependencies**: Imports from root `core/` by appending the root directory to `sys.path`.
-   **Functionality**: Handles file uploads, processes them asynchronously, and stores results (currently integrated with Supabase logic).

### 4. Desktop Application (`desktop/`)
-   **Entry Point**: `desktop/GUI_pyside.py`
-   **Framework**: PySide6 (Qt) with `QWebEngineView` for the UI.
-   **Dependencies**: **Crucial Note:** The desktop application currently uses a **local copy** of the core library located in `desktop/core/` and `desktop/style_guides.py`.
    -   This design decision ensures the desktop application is self-contained and easier to package (e.g., with PyInstaller) without complex path manipulation.
    -   **Developer Note**: If you modify logic in the root `core/`, you must verify if it needs to be propagated to `desktop/core/` to maintain feature parity.

## 💡 Key Concepts

### Document Structure Detection
Formatly does not use regex for everything. It employs a two-step process:
1.  **AI Analysis**: The document text is sent to an LLM (Gemini) with a prompt to return a JSON structure representing the "blocks" of the document (e.g., "This paragraph is a Heading 1", "This is a list item").
2.  **Deterministic Formatting**: The `AdvancedFormatter` parses this JSON and applies standard Word styles to the underlying XML of the `.docx` file using `python-docx`.

### Style Guides
Styles are defined in `style_guides.py` as pure Python dictionaries. This makes it easy to tweak font sizes or add new styles without changing the core logic.

## ⚠️ Known Issues / Technical Debt

-   **Code Duplication**: As mentioned, `desktop/core` is a duplicate of root `core`. A future refactor should aim to package `core` as an installable Python package that both the CLI and Desktop app can import as a dependency.
