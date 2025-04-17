#!/usr/bin/env python3
"""
DOCX Tool - Command Line Interface
---------------------------------
A simple command-line tool for processing DOCX documents:
1. Basic text replacement
2. ELISA document conversion

Usage:
    python docx_tool_cli.py replace <input_file> <output_file> <find_text> <replace_text>
    python docx_tool_cli.py replace_multiple <input_file> <output_file>
    python docx_tool_cli.py elisa <outside_doc> <template_doc> <output_file> <catalog_no> <lot_no>
"""

import sys
import os
import argparse
from document_processor import process_document
from elisa_document_converter import process_elisa_document

def print_header():
    """Print the tool header."""
    print("\n" + "=" * 80)
    print("DOCX Tool - Command Line Interface".center(80))
    print("=" * 80)

def replace_text(input_file, output_file, find_text, replace_text):
    """Replace text in a DOCX file."""
    print(f"\nReplacing '{find_text}' with '{replace_text}' in {input_file}...")
    
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found.")
        return
    
    try:
        count = process_document(input_file, output_file, find_text, replace_text)
        print(f"\nSuccess! Made {count} replacements.")
        print(f"Output saved to: {output_file}")
    except Exception as e:
        print(f"Error processing document: {str(e)}")

def replace_multiple_text(input_file, output_file):
    """Replace multiple text patterns in a DOCX file."""
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' not found.")
        return
    
    replacements = []
    print("\nEnter text replacements (leave both fields empty to finish):")
    
    while True:
        find_text = input("\nText to find: ").strip()
        if not find_text:
            break
            
        replace_text = input("Replace with (leave empty to delete): ").strip()
        replacements.append((find_text, replace_text))
        
        print(f"Added: '{find_text}' -> '{replace_text}'")
        
    if not replacements:
        print("No replacements specified. Exiting.")
        return
    
    print(f"\nProcessing document with {len(replacements)} replacements...")
    
    try:
        current_input = input_file
        temp_output = output_file
        total_count = 0
        
        for i, (find_text, replace_text) in enumerate(replacements):
            if i > 0:
                current_input = temp_output
                
            count = process_document(current_input, temp_output, find_text, replace_text)
            total_count += count
            print(f"- Replacement {i+1}: '{find_text}' -> '{replace_text}', {count} occurrences")
            
        print(f"\nSuccess! Made {total_count} replacements across {len(replacements)} search terms.")
        print(f"Output saved to: {output_file}")
    except Exception as e:
        print(f"Error processing document: {str(e)}")

def convert_elisa_document(outside_doc, template_doc, output_file, catalog_no, lot_no):
    """Convert an ELISA document."""
    print(f"\nConverting ELISA document...")
    print(f"Outside document: {outside_doc}")
    print(f"Template document: {template_doc}")
    print(f"Catalog #: {catalog_no}")
    print(f"Lot #: {lot_no}")
    
    if not os.path.exists(outside_doc):
        print(f"Error: Outside document '{outside_doc}' not found.")
        return
        
    if not os.path.exists(template_doc):
        print(f"Error: Template document '{template_doc}' not found.")
        return
    
    try:
        generated_path = process_elisa_document(outside_doc, template_doc, catalog_no, lot_no)
        
        # If output file is specified and different from generated path, copy the file
        if output_file and output_file != generated_path:
            import shutil
            shutil.copy2(generated_path, output_file)
            print(f"\nSuccess! ELISA document converted.")
            print(f"Output saved to: {output_file}")
        else:
            print(f"\nSuccess! ELISA document converted.")
            print(f"Output saved to: {generated_path}")
    except Exception as e:
        print(f"Error converting ELISA document: {str(e)}")
        print("Check that both documents are valid DOCX files.")

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description="DOCX Tool - Process DOCX documents")
    subparsers = parser.add_subparsers(dest="command", help="Command to execute")
    
    # Replace text command
    replace_parser = subparsers.add_parser("replace", help="Replace text in a DOCX document")
    replace_parser.add_argument("input_file", help="Input DOCX file")
    replace_parser.add_argument("output_file", help="Output DOCX file")
    replace_parser.add_argument("find_text", help="Text to find")
    replace_parser.add_argument("replace_text", help="Text to replace with (leave empty to delete)", nargs="?", default="")
    
    # Replace multiple text command
    replace_multiple_parser = subparsers.add_parser("replace_multiple", help="Replace multiple text patterns in a DOCX document")
    replace_multiple_parser.add_argument("input_file", help="Input DOCX file")
    replace_multiple_parser.add_argument("output_file", help="Output DOCX file")
    
    # ELISA document conversion command
    elisa_parser = subparsers.add_parser("elisa", help="Convert an ELISA document")
    elisa_parser.add_argument("outside_doc", help="Outside document (Boster ELISA documentation)")
    elisa_parser.add_argument("template_doc", help="Template document (Innovative Research template)")
    elisa_parser.add_argument("output_file", help="Output DOCX file")
    elisa_parser.add_argument("catalog_no", help="Catalog number")
    elisa_parser.add_argument("lot_no", help="Lot number")
    
    args = parser.parse_args()
    
    print_header()
    
    if args.command == "replace":
        replace_text(args.input_file, args.output_file, args.find_text, args.replace_text)
    elif args.command == "replace_multiple":
        replace_multiple_text(args.input_file, args.output_file)
    elif args.command == "elisa":
        convert_elisa_document(args.outside_doc, args.template_doc, args.output_file, args.catalog_no, args.lot_no)
    else:
        parser.print_help()
        print("\nExamples:")
        print("  python docx_tool_cli.py replace input.docx output.docx \"old text\" \"new text\"")
        print("  python docx_tool_cli.py replace_multiple input.docx output.docx")
        print("  python docx_tool_cli.py elisa outside.docx template.docx output.docx \"CATALOG123\" \"LOT456\"")

if __name__ == "__main__":
    main()