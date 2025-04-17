# DOCX Text Editor & ELISA Document Converter

A document processing tool with multiple functions:
1. Replace or delete text in DOCX documents
2. Convert specialized ELISA documentation from Boster to Innovative Research format

## Installation

### Prerequisites
- Python 3.10 or later
- pip (Python package manager)

### Setup

1. Clone or download this repository
2. Install the required dependencies:

```bash
pip install flask docx flask-sqlalchemy gunicorn python-docx email-validator werkzeug
```

## Usage

### Command Line Interface

A command-line interface is available for direct document processing without using the web interface.

#### Basic Text Replacement

```bash
python docx_tool_cli.py replace input.docx output.docx "text to find" "replacement text"
```

To delete text, leave the "replacement text" empty:

```bash
python docx_tool_cli.py replace input.docx output.docx "text to delete" ""
```

#### Multiple Text Replacements

```bash
python docx_tool_cli.py replace_multiple input.docx output.docx
```

This will prompt you to enter multiple find/replace pairs interactively.

#### ELISA Document Conversion

```bash
python docx_tool_cli.py elisa outside_document.docx template_document.docx output.docx "CATALOG123" "LOT456"
```

Where:
- `outside_document.docx` is the Boster ELISA kit documentation
- `template_document.docx` is the Innovative Research template
- `CATALOG123` is the catalog number to insert
- `LOT456` is the lot number to insert

### Web Interface

To run the web application:

```bash
python main.py
```

Then open your browser and navigate to: http://localhost:5000

The web interface offers:
1. Drag-and-drop file upload
2. Multiple text replacements in a single operation
3. Interactive ELISA document converter
4. Step-by-step tutorial for new users

## ELISA Document Converter

The ELISA Document Converter is a specialized tool that:

1. Extracts specific sections from Boster ELISA kit documentation
2. Reorganizes content into an Innovative Research template
3. Replaces branding (Boster → Innovative Research, removes PicoKine)
4. Formats text according to specifications:
   - Calibri 11pt for body text
   - Calibri 12pt bold blue for headings
   - 1.15 line spacing
5. Preserves tables and images
6. Inserts catalog and lot numbers
7. Adds required disclaimer text

## Files Overview

- `main.py` - Entry point for the web application
- `app.py` - Main Flask application
- `document_processor.py` - Basic text replacement functionality
- `elisa_document_converter.py` - ELISA document conversion logic
- `docx_tool_cli.py` - Command-line interface
- `templates/` - HTML templates for the web interface
- `static/` - CSS, JS, and other static files