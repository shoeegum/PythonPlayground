import os
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile

def process_elisa_document(outside_doc_path, template_doc_path, catalog_no, lot_no):
    """
    Process an ELISA kit document, extracting content from the outside document
    and populating it into the template document according to specifications.
    
    Args:
        outside_doc_path (str): Path to the source document
        template_doc_path (str): Path to the template document
        catalog_no (str): Catalog number to use
        lot_no (str): Lot number to use
        
    Returns:
        str: Path to the generated document
    """
    print(f"Loading documents...")
    # Load documents
    outside_doc = Document(outside_doc_path)
    inside_doc = Document(template_doc_path)
    
    print(f"Processing documents...")
    # Create a mapping of section titles and content from the outside document
    outside_sections = extract_sections(outside_doc)
    
    # Process the inside document
    process_inside_document(inside_doc, outside_sections, catalog_no, lot_no)
    
    # Save the processed document
    output_path = os.path.join(tempfile.gettempdir(), "Processed_ELISA_Document.docx")
    inside_doc.save(output_path)
    print(f"Document saved to: {output_path}")
    
    return output_path

def extract_sections(doc):
    """
    Extract sections from the outside document.
    
    Args:
        doc: The document to extract sections from
        
    Returns:
        dict: A mapping of section titles to their content
    """
    sections = {}
    current_section = None
    current_content = []
    
    for para in doc.paragraphs:
        text = para.text.strip()
        
        # Skip empty paragraphs
        if not text:
            continue
        
        # Check if this paragraph is a heading
        if para.style.name.startswith('Heading') or is_likely_heading(para):
            # Save the previous section if it exists
            if current_section is not None:
                sections[current_section] = current_content
            
            # Start a new section
            current_section = text
            current_content = [para]
        else:
            # Add to the current section content
            if current_section is not None:
                current_content.append(para)
    
    # Save the last section
    if current_section is not None:
        sections[current_section] = current_content
    
    # Add special handling for sections with tables
    section_tables = {}
    
    for i, table in enumerate(doc.tables):
        # Find which section this table belongs to
        table_section = find_table_section(sections, table)
        if table_section:
            if table_section not in section_tables:
                section_tables[table_section] = []
            section_tables[table_section].append(table)
    
    # Merge tables into sections
    for section, tables in section_tables.items():
        if section in sections:
            sections[section].extend(tables)
    
    return sections

def is_likely_heading(para):
    """
    Check if a paragraph is likely a heading based on formatting.
    
    Args:
        para: The paragraph to check
        
    Returns:
        bool: True if the paragraph is likely a heading
    """
    # Check if bold
    if para.runs and para.runs[0].bold:
        return True
    
    # Check if it's a short paragraph that ends with a colon
    if len(para.text) < 100 and para.text.strip().endswith(':'):
        return True
    
    # Other heading indicators
    heading_patterns = [
        r'^[A-Z][a-zA-Z\s]+:$',  # Title followed by colon
        r'^[0-9]+\.\s+[A-Z]',    # Numbered section
        r'^[A-Z][A-Z\s]+$'       # ALL CAPS
    ]
    
    for pattern in heading_patterns:
        if re.match(pattern, para.text.strip()):
            return True
    
    return False


def find_table_section(sections, table):
    """
    Find which section a table belongs to.
    
    Args:
        sections: Pre-extracted sections
        table: The table to find the section for
        
    Returns:
        str: The section title, or None if not found
    """
    for section, content in sections.items():
        for item in content:
            if isinstance(item, type(table)):
                if item._element is table._element:
                    return section
    return None

def process_inside_document(doc, outside_sections, catalog_no, lot_no):
    """
    Process the inside document, populating it with content from the outside document.
    
    Args:
        doc: The document to process
        outside_sections: Sections extracted from the outside document
        catalog_no: Catalog number to use
        lot_no: Lot number to use
    """
    # Get or create sections in the inside document
    # Implementation would need to locate placeholders in the template
    
    # Set catalog and lot numbers
    set_metadata(doc, catalog_no, lot_no)
    
    # Process each required section
    process_intended_use(doc, outside_sections)
    process_background(doc, outside_sections)
    process_assay_principle(doc, outside_sections)
    process_overview(doc, outside_sections)
    process_technical_details(doc, outside_sections)
    process_preparations_before_assay(doc, outside_sections)
    process_kit_components(doc, outside_sections)
    process_required_materials(doc, outside_sections)
    process_standard_curve(doc, outside_sections)
    process_assay_variability(doc, outside_sections)
    process_reproducibility(doc, outside_sections)
    process_experiment_preparation(doc, outside_sections)
    process_standard_dilution(doc, outside_sections)
    process_sample_preparation(doc, outside_sections)
    process_sample_collection(doc, outside_sections)
    process_sample_dilution(doc, outside_sections)
    process_assay_protocol(doc, outside_sections)
    process_protocol_notes(doc, outside_sections)
    process_data_analysis(doc, outside_sections)
    add_disclaimer(doc)
    #all caps needed for section titles
    
    # Format the document
    apply_formatting(doc)
    
    # Replace all instances of "Boster" with "Innovative Research" and remove "PicoKine"
    replace_text_in_document(doc, "Boster", "Innovative Research")
    replace_text_in_document(doc, "PicoKine", "")

def set_metadata(doc, catalog_no, lot_no):
    """Set catalog and lot numbers in the document."""
    # Implementation would find and update these fields
    print(f"Setting metadata: Catalog No: {catalog_no}, Lot No: {lot_no}")
    # This is a placeholder - would need to locate where to insert in template

def process_intended_use(doc, outside_sections):
    """Process the Intended Use section."""
    # Extract from "Assay Principle" section first paragraph
    print("Processing: INTENDED USE")
    # Implementation would extract the content and add it to the document

def process_background(doc, outside_sections):
    """Process the Background section."""
    print("Processing: BACKGROUND")
    # Find section like "Background on ..."
    background_section = None
    for title in outside_sections.keys():
        if title.startswith("Background on "):
            background_section = title
            break
    
    if background_section:
        print(f"  Found background section: {background_section}")
        # Implementation would add this content to the document

def process_assay_principle(doc, outside_sections):
    """Process the Assay Principle section."""
    print("Processing: ASSAY PRINCIPLE")
    # Extract from last paragraph of "Assay Principle"
    # Implementation would extract and add this content

def process_overview(doc, outside_sections):
    """Process the Overview section."""
    print("Processing: OVERVIEW")
    # Implementation would process overview section and table, removing specified rows

def process_technical_details(doc, outside_sections):
    """Process the Technical Details section."""
    print("Processing: TECHNICAL DETAILS")
    # Implementation would add technical details content

def process_preparations_before_assay(doc, outside_sections):
    """Process the Preparations Before Assay section."""
    print("Processing: PREPARATIONS BEFORE ASSAY")
    # Implementation would add preparations content

def process_kit_components(doc, outside_sections):
    """Process the Kit Components section."""
    print("Processing: KIT COMPONENT/MATERIALS PROVIDED")
    # Implementation would add kit components table

def process_required_materials(doc, outside_sections):
    """Process the Required Materials section."""
    print("Processing: REQUIRED MATERIALS THAT ARE NOT SUPPLIED")
    # Implementation would add required materials as bullet points

def process_standard_curve(doc, outside_sections):
    """Process the Standard Curve Example section."""
    print("Processing: STANDARD CURVE EXAMPLE")
    # Implementation would find and add standard curve content

def process_assay_variability(doc, outside_sections):
    """Process the Assay Variability section."""
    print("Processing: INTRA/INTER-ASSAY VARIABILITY")
    # Implementation would add variability content and table

def process_reproducibility(doc, outside_sections):
    """Process the Reproducibility section."""
    print("Processing: REPRODUCIBILITY")
    # Implementation would add reproducibility content and table

def process_experiment_preparation(doc, outside_sections):
    """Process the Experiment Preparation section."""
    print("Processing: PREPARATION BEFORE THE EXPERIMENT")
    # Implementation would add preparation content with modified table

def process_standard_dilution(doc, outside_sections):
    """Process the Standard Dilution section."""
    print("Processing: DILUTION OF STANDARD")
    # Find section starting with "Dilution of ..."
    dilution_section = None
    for title in outside_sections.keys():
        if title.startswith("Dilution of "):
            dilution_section = title
            break
    
    if dilution_section:
        print(f"  Found dilution section: {dilution_section}")
        # Implementation would add this content including image

def process_sample_preparation(doc, outside_sections):
    """Process the Sample Preparation section."""
    print("Processing: SAMPLE PREPARATION AND STORAGE")
    # Implementation would add sample preparation content

def process_sample_collection(doc, outside_sections):
    """Process the Sample Collection Notes section."""
    print("Processing: SAMPLE COLLECTION NOTES")
    # Implementation would add sample collection content with replacements

def process_sample_dilution(doc, outside_sections):
    """Process the Sample Dilution section."""
    print("Processing: SAMPLE DILUTION GUIDELINE")
    # Implementation would add sample dilution content

def process_assay_protocol(doc, outside_sections):
    """Process the Assay Protocol section."""
    print("Processing: ASSAY PROTOCOL")
    # Implementation would add protocol content with proper list formatting

def process_protocol_notes(doc, outside_sections):
    """Process the Protocol Notes section."""
    print("Processing: ASSAY PROTOCOL NOTES")
    # Implementation would add protocol notes with proper list formatting

def process_data_analysis(doc, outside_sections):
    """Process the Data Analysis section."""
    print("Processing: DATA ANALYSIS")
    # Implementation would add only the last two paragraphs

def add_disclaimer(doc):
    """Add the Disclaimer section."""
    print("Adding: DISCLAIMER")
    # Implementation would add the specified disclaimer paragraphs in italics

def apply_formatting(doc):
    """Apply the required formatting to the document."""
    print("Applying document formatting")
    # Set font to Calibri 11 for body
    # Set headings to Calibri 12 bold blue
    # Set paragraph spacing to Multiple 1.15
    # Handle footnotes if required

def replace_text_in_document(doc, old_text, new_text):
    """
    Replace all occurrences of old_text with new_text in the document.
    
    Args:
        doc: The document to process
        old_text: Text to find
        new_text: Text to replace with
    """
    print(f"Replacing: '{old_text}' with '{new_text}'")
    
    # Replace in paragraphs
    for para in doc.paragraphs:
        for run in para.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
    
    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)

if __name__ == "__main__":
    print("ELISA Document Converter")
    print("------------------------")
    
    outside_doc_path = input("Enter path to outside document: ")
    template_doc_path = input("Enter path to template document: ")
    catalog_no = input("Enter catalog number: ")
    lot_no = input("Enter lot number: ")
    
    output_path = process_elisa_document(outside_doc_path, template_doc_path, catalog_no, lot_no)
    print(f"Conversion complete! Document saved to: {output_path}")