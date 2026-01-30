<<<<<<< HEAD
# This readme is not updated
# Formatly 🎓

An AI-powered academic document formatter that automatically applies proper citation styles (APA, MLA, Chicago) to Word documents using Google's Gemini AI.

## ✨ Features

- **🤖 AI-Powered Document Analysis**: Uses Google Gemini AI to intelligently detect document structure
- **📚 Multiple Citation Styles**: Supports APA, MLA, and Chicago formatting styles
- **🎯 Smart Formatting**: Automatically applies proper fonts, margins, spacing, and heading styles
- **📝 Reference Formatting**: Intelligent formatting of citations and references
- **🔤 Advanced Spelling \u0026 Grammar Check**: Built-in spell checking with AI-powered suggestions and quality scoring
- **⚡ Dynamic Chunking**: Intelligent text processing with optimal chunk size calculation for efficient API usage
- **🛡️ Robust Rate Limit Management**: Automatic detection and handling of API quota limits with intelligent retry logic
- **🔄 Smart Error Recovery**: Graceful handling of temporary API errors with exponential backoff
- **💻 User-Friendly**: Interactive mode, command-line interface, and comprehensive error handling
- **🔒 Secure**: Environment variable-based API key management

## 🚀 Installation

### Prerequisites

- Python 3.7 or higher
- Word documents (.docx format)
- Google Gemini API key

### Setup

1. **Clone or download the project:**
   ```bash
   git clone <repository-url>
   cd "Formatly V3"
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up your API key:**
   - Get a Google Gemini API key from [Google AI Studio](https://makersuite.google.com/app/apikey)
   - Create a `.env` file in the project directory:
     ```
     GEMINI_API_KEY=your_api_key_here
     GEMINI_MODEL=gemini-2.0-flash-exp
     ```

## 📖 Usage

### Command Line Interface

**Basic usage:**
```bash
python app.py document.docx
```

**Specify output file:**
```bash
python app.py document.docx -o formatted_document.docx
```

**Choose citation style:**
```bash
python app.py document.docx --style mla
python app.py document.docx --style chicago
python app.py document.docx --style apa  # default
```

**Interactive mode:**
```bash
python app.py --interactive
```

**List available styles:**
```bash
python app.py --list-styles
```

**Check spelling and grammar only:**
```bash
python app.py document.docx --spell-only
```

**Check spelling before formatting:**
```bash
python app.py document.docx --check-spelling
```

### Command Line Options

```
positional arguments:
  input                 Input Word document (required if not using --interactive)

optional arguments:
  -h, --help            show this help message and exit
  -o, --output OUTPUT   Output file path (default: <input>_formatted_<style>.docx)
  -s, --style {apa,mla,chicago}
                        Citation style: apa, mla, or chicago (default: apa)
  -i, --interactive     Interactive mode - prompts for input file if not provided
  -l, --list-styles     List available citation styles and exit
  -c, --check-spelling  Check spelling and grammar before formatting
  --spell-only          Only check spelling/grammar, don't format document
```

## 🎨 Supported Citation Styles

### APA Style
- **Title page:** Yes
- **Abstract:** Required
- **Font:** Times New Roman, 12pt
- **Margins:** 1 inch all around
- **Line spacing:** Double
- **Page numbers:** Top right header

### MLA Style
- **Title page:** No
- **Abstract:** Optional
- **Font:** Times New Roman, 12pt
- **Margins:** 1 inch all around
- **Line spacing:** Double
- **Page numbers:** Top right header

### Chicago Style
- **Title page:** Yes
- **Abstract:** Optional
- **Font:** Times New Roman, 12pt
- **Margins:** 1 inch all around
- **Line spacing:** Double
- **Page numbers:** Top right header

## 🔧 Configuration

### Environment Variables

Create a `.env` file in your project directory:

```env
# Required
GEMINI_API_KEY=your_gemini_api_key_here

# Optional
GEMINI_MODEL=gemini-2.0-flash-exp
MAX_RETRIES=3
TIMEOUT=30
LOG_LEVEL=INFO
```

### Configuration Options

- **GEMINI_API_KEY**: Your Google Gemini API key (required)
- **GEMINI_MODEL**: The Gemini model to use (default: gemini-2.0-flash-exp)
- **MAX_RETRIES**: Number of retry attempts for API calls (default: 3)
- **TIMEOUT**: Request timeout in seconds (default: 30)
- **LOG_LEVEL**: Logging level (default: INFO)

### Advanced Features

#### 🛡️ Rate Limit Management

Formatly V6 includes intelligent rate limit handling that automatically:

- **Detects API quota limits**: Automatically identifies different types of rate limits (RPM, RPD, TPM)
- **Smart retry logic**: Implements exponential backoff with jitter for temporary limits
- **Graceful degradation**: Stops processing when daily limits are reached with clear user feedback
- **Intelligent wait times**: Calculates optimal wait times based on API response headers

**Rate Limit Types Handled:**
- **RPM (Requests Per Minute)**: Automatic retry with calculated delay
- **RPD (Requests Per Day)**: Graceful shutdown with upgrade suggestions
- **TPM (Tokens Per Minute)**: Dynamic adjustment of processing speed

#### ⚡ Dynamic Chunking System

The application uses intelligent chunk sizing to optimize API usage:

- **Adaptive chunk sizes**: Automatically adjusts based on document complexity
- **Rate limit awareness**: Reduces chunk size when approaching limits
- **Performance optimization**: Maximizes throughput while respecting quotas
- **Error recovery**: Automatically retries failed chunks with adjusted sizing

**Chunk Calculation Factors:**
- Document length and complexity
- Current API rate limit status
- Historical processing performance
- Available token budget

## 📁 Project Structure

```
Formatly V3/
├── app.py              # Main application file
├── style_guides.py     # Citation style definitions
├── config.py           # Configuration management
├── requirements.txt    # Python dependencies
├── .env               # Environment variables (create this)
├── README.md          # This file
└── *.docx             # Sample documents
```

## 🤝 How It Works

1. **Document Analysis**: The application extracts text from your Word document
2. **AI Processing**: Sends content to Google Gemini AI to identify document structure
3. **Style Application**: Applies the selected citation style formatting
4. **Output Generation**: Saves the formatted document with proper styling

## 🔍 Examples

### Example 1: Format with APA style
```bash
python app.py my_paper.docx --style apa --output apa_formatted.docx
```

### Example 2: Interactive mode
```bash
python app.py --interactive
# Follow the prompts to select your document and style
```

### Example 3: Check available styles
```bash
python app.py --list-styles
```

## 🛠️ Troubleshooting

### Common Issues

**Error: "GEMINI_API_KEY environment variable not set"**
- Solution: Create a `.env` file with your API key

**Error: "Input file must be a Word document (.docx)"**
- Solution: Ensure your file has a `.docx` extension

**Error: "Permission denied when accessing file"**
- Solution: Close the document if it's open in Word

**Error: "Input file not found"**
- Solution: Check the file path and ensure the file exists

### API Key Issues

1. **Get your API key:**
   - Visit [Google AI Studio](https://makersuite.google.com/app/apikey)
   - Create a new API key
   - Add it to your `.env` file

2. **Verify API key:**
   ```bash
   python app.py --list-styles
   ```

## 📝 Requirements

- Python 3.7+
- google-generativeai
- python-docx
- python-dotenv

## 🔒 Security

- API keys are stored securely in environment variables
- No sensitive data is logged or stored
- All API communications are encrypted

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

If you encounter any issues:

1. Check this README for common solutions
2. Verify your API key is correctly set
3. Ensure your document is in .docx format
4. Check that the file is not currently open in Word

## 🎯 Version History

- **v3.0.0**: Major rewrite with AI integration, multiple citation styles, and improved error handling
- **v2.x**: Previous versions with basic formatting
- **v1.x**: Initial release

---

**Made with ❤️ for academic writers everywhere**
=======
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
>>>>>>> df2a4285aed60fd8df3b0d2ce247c5ac4431d138
