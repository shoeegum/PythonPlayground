# DOCX Text Editor & ELISA Document Converter (Standalone Version)

A command-line tool for processing DOCX documents without requiring the web interface.

## Files You Need

To run this tool locally, you need these files:
1. `document_processor.py` - Handles basic text replacement
2. `elisa_document_converter.py` - ELISA document converter functionality
3. `docx_tool_cli.py` - Command-line interface

## Installation

### Prerequisites

1. **Python 3.10 or later** - Download from https://www.python.org/downloads/
2. **python-docx library** - Install with:
   ```
   pip install python-docx
   ```

### Setup Instructions

1. Create a new folder on your computer (e.g., "docx_tool")
2. Download the following files from this Replit project:
   - `document_processor.py`
   - `elisa_document_converter.py`
   - `docx_tool_cli.py`
3. Place all three files in your folder

## Usage

Open Command Prompt (Windows) or Terminal (Mac/Linux), navigate to your folder, and run:

### Basic Text Replacement

```bash
python docx_tool_cli.py replace input.docx output.docx "text to find" "replacement text"
```

### Delete Text

```bash
python docx_tool_cli.py replace input.docx output.docx "text to delete" ""
```

### Multiple Text Replacements

```bash
python docx_tool_cli.py replace_multiple input.docx output.docx
```

This will prompt you to enter multiple find/replace pairs interactively.

### ELISA Document Conversion

```bash
python docx_tool_cli.py elisa outside_document.docx template_document.docx output.docx "CATALOG123" "LOT456"
```

Where:
- `outside_document.docx` is the Boster ELISA kit documentation
- `template_document.docx` is the Innovative Research template
- `output.docx` is where to save the result
- `CATALOG123` is the catalog number to insert
- `LOT456` is the lot number to insert

## Troubleshooting

1. **File Not Found Errors**: Make sure the input documents exist in the current directory or provide full file paths
2. **Import Errors**: Ensure python-docx is installed (`pip install python-docx`)
3. **Permission Errors**: Make sure you have permission to read/write in the specified locations