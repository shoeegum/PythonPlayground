import os
import logging
import tempfile
import uuid
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
from werkzeug.utils import secure_filename
from document_processor import process_document
from elisa_document_converter import process_elisa_document

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-key-for-testing")
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()  # Use temp directory for uploads
app.config['ALLOWED_EXTENSIONS'] = {'docx'}

def allowed_file(filename):
    """Check if the file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/', methods=['GET'])
def index():
    """Render the main page."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and text replacement."""
    # Check if a file was uploaded
    if 'document' not in request.files:
        flash('No file part', 'danger')
        return redirect(url_for('index'))
    
    file = request.files['document']
    
    # Check if the file is selected
    if file.filename == '':
        flash('No file selected', 'danger')
        return redirect(url_for('index'))
    
    # Check if file type is allowed
    if not allowed_file(file.filename):
        flash('Invalid file type. Only DOCX files are allowed.', 'danger')
        return redirect(url_for('index'))
    
    # Get text replacement parameters
    # Check if we have multiple replacements
    find_texts = request.form.getlist('find_text[]')
    replace_texts = request.form.getlist('replace_text[]')
    
    # If no multiple replacements were submitted, fall back to single replacement fields
    if not find_texts:
        find_text = request.form.get('find_text', '').strip()
        replace_text = request.form.get('replace_text', '').strip()
        
        if not find_text:
            flash('Please enter at least one text to find', 'danger')
            return redirect(url_for('index'))
        
        find_texts = [find_text]
        replace_texts = [replace_text]
    
    # Filter out empty find_text entries
    replacements = []
    for i, find_text in enumerate(find_texts):
        find_text = find_text.strip()
        if find_text:
            replace_text = replace_texts[i].strip() if i < len(replace_texts) else ''
            replacements.append((find_text, replace_text))
    
    if not replacements:
        flash('Please enter at least one text to find', 'danger')
        return redirect(url_for('index'))
    
    try:
        # Create a unique filename to avoid collisions
        orig_filename = secure_filename(file.filename)
        base_filename = str(uuid.uuid4())
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{base_filename}_input.docx")
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{base_filename}_output.docx")
        
        # Save the uploaded file
        file.save(input_path)
        
        # Process the document with all replacements
        total_count = 0
        current_input = input_path
        
        for i, (find_text, replace_text) in enumerate(replacements):
            # For the first replacement, use the original input file
            # For subsequent replacements, use the output of the previous replacement
            if i > 0:
                current_input = output_path
            
            count = process_document(current_input, output_path, find_text, replace_text)
            total_count += count
            
            logger.debug(f"Replacement {i+1}: '{find_text}' -> '{replace_text}', count: {count}")
        
        if total_count > 0:
            flash(f'Successfully made {total_count} replacements across {len(replacements)} search terms', 'success')
        else:
            flash('No occurrences found for any of the search terms', 'info')
        
        # Store the output path and original filename in session for download
        session['output_path'] = output_path
        session['download_filename'] = f"modified_{orig_filename}"
        
        return redirect(url_for('index'))

    
    except Exception as e:
        logger.exception("Error processing document")
        flash(f'Error processing document: {str(e)}', 'danger')
        return redirect(url_for('index'))

@app.route('/elisa', methods=['GET', 'POST'])
def elisa_converter():
    """Handle ELISA converter page and document conversion."""
    if request.method == 'GET':
        return render_template('elisa_converter.html')
    """Handle ELISA document conversion."""
    # Check if the required files were uploaded
    if 'outside_document' not in request.files or 'template_document' not in request.files:
        flash('Missing document files', 'danger')
        return redirect(url_for('elisa_converter'))
    
    outside_file = request.files['outside_document']
    template_file = request.files['template_document']
    
    # Check if the files are selected
    if outside_file.filename == '' or template_file.filename == '':
        flash('Please select both outside document and template document', 'danger')
        return redirect(url_for('elisa_converter'))
    
    # Check if file types are allowed
    if not allowed_file(outside_file.filename) or not allowed_file(template_file.filename):
        flash('Invalid file type. Only DOCX files are allowed.', 'danger')
        return redirect(url_for('elisa_converter'))
    
    # Get catalog and lot numbers
    catalog_no = request.form.get('catalog_no', '').strip()
    lot_no = request.form.get('lot_no', '').strip()
    
    if not catalog_no or not lot_no:
        flash('Please provide both catalog number and lot number', 'danger')
        return redirect(url_for('elisa_converter'))
    
    try:
        # Create unique filenames
        outside_filename = f"{uuid.uuid4()}_outside.docx"
        template_filename = f"{uuid.uuid4()}_template.docx"
        
        outside_path = os.path.join(app.config['UPLOAD_FOLDER'], outside_filename)
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_filename)
        
        # Save the uploaded files
        outside_file.save(outside_path)
        template_file.save(template_path)
        
        # Process the ELISA document
        logger.debug(f"Processing ELISA document: {outside_path} with template {template_path}")
        output_path = process_elisa_document(outside_path, template_path, catalog_no, lot_no)
        
        # Store the output path for download
        session['output_path'] = output_path
        session['download_filename'] = f"ELISA_Document_{catalog_no}.docx"
        
        flash('Document conversion completed successfully', 'success')
        return redirect(url_for('elisa_converter', processed=True))
    
    except Exception as e:
        logger.exception("Error converting ELISA document")
        flash(f'Error converting document: {str(e)}', 'danger')
        return redirect(url_for('elisa_converter'))

@app.route('/download', methods=['GET'])
def download_file():
    """Download the processed file."""
    output_path = session.get('output_path')
    download_filename = session.get('download_filename')
    
    if not output_path or not download_filename or not os.path.exists(output_path):
        flash('No processed file available for download', 'danger')
        return redirect(url_for('index'))
    
    try:
        return send_file(
            output_path,
            as_attachment=True,
            download_name=download_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        logger.exception("Error downloading file")
        flash(f'Error downloading file: {str(e)}', 'danger')
        return redirect(url_for('index'))

@app.errorhandler(413)
def request_entity_too_large(error):
    """Handle file size too large errors."""
    flash('File too large. Maximum size is 16MB.', 'danger')
    return redirect(url_for('index'))

@app.errorhandler(500)
def internal_server_error(error):
    """Handle internal server errors."""
    flash('An unexpected error occurred. Please try again later.', 'danger')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
