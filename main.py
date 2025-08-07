from flask import Flask, render_template, request, redirect, url_for, flash
import os
import logging
from werkzeug.utils import secure_filename
import zipfile
import shutil
import tempfile
from pathlib import Path
import re
import json

# Initialize Flask application
app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Change this to a secure random key

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configure upload folder and allowed extensions
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'zip'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Define a global variable for tracking processing status
processing = False

# Utility Functions
def allowed_file(filename):
    """Check if the file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_video_position_params(form):
    """Extract and convert video position parameters from form data."""
    try:
        start_time = int(form.get('startTime', 0))
        end_time = int(form.get('endTime', 0))
        frequency = int(form.get('frequency', 1))
        return {'start_time': start_time, 'end_time': end_time, 'frequency': frequency}
    except ValueError as e:
        logger.error(f"Error converting video position parameters: {e}")
        return None

def process_uploaded_file(file, annotation_option='with_annos', is_multi_tab=True, implement_video_frames=False, video_position_params=None):
    """
    Process the uploaded file: unzip, organize, and prepare data for rendering.
    Now includes options for handling annotation and video frames.
    """
    global processing
    if processing:
        return {'success': False, 'message': 'A file is already being processed. Please wait.'}
    processing = True  # Set the flag to indicate processing is in progress

    try:
        from utils.presentation_generator import PresentationGenerator
        
        # Extract original filename without extension
        original_filename = os.path.splitext(file.filename)[0]
        
        # Create a temporary directory to store the contents of the zip file
        temp_dir = tempfile.mkdtemp()
        zip_path = os.path.join(temp_dir, secure_filename(file.filename))
        file.save(zip_path)

        # Extract the zip file
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Generate the presentation using PresentationGenerator
        generator = PresentationGenerator()
        
        # Create outputs directory if it doesn't exist
        outputs_dir = 'outputs'
        os.makedirs(outputs_dir, exist_ok=True)
        
        # Generate presentation
        ppt_filename = generator.generate_from_folder(
            temp_dir, 
            annotation_option=annotation_option,
            implement_video_frames=implement_video_frames,
            original_filename=original_filename
        )
        
        # Get actual slide count from the generated presentation
        from pptx import Presentation
        prs = Presentation(ppt_filename)
        slide_count = len(prs.slides)
        
        # Get folder count and other info
        folder_structure = generator._organize_folder_structure(temp_dir)
        folder_count = len(folder_structure)
        
        # Check if video folder exists
        video_folder_found = any('video' in folder.lower() for folder in folder_structure.keys())

        # Clean up: Remove the temporary directory and its contents
        shutil.rmtree(temp_dir)

        return {
            'success': True,
            'filename': file.filename,
            'ppt_file': os.path.basename(ppt_filename),
            'folder_count': folder_count,
            'slide_count': slide_count,
            'video_folder_found': video_folder_found,
        }

    except Exception as e:
        logger.error(f"Error processing file: {str(e)}")
        return {'success': False, 'message': str(e)}
    finally:
        processing = False  # Reset the processing flag

# Flask Routes
@app.route('/', methods=['GET'])
def index():
    """Render the index page with file upload form."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and processing."""
    try:
        if 'file' not in request.files:
            flash('No file selected', 'error')
            return redirect(url_for('index'))

        file = request.files['file']
        if file.filename == '':
            flash('No file selected', 'error')
            return redirect(url_for('index'))

        # Get form parameters
        annotation_option = request.form.get('annotation_option', 'with_annos')
        implement_video_frames = request.form.get('implementVideoFrames') == 'true'

        # Get video position parameters if video frames are enabled
        video_position_params = None
        if implement_video_frames:
            video_position_params = extract_video_position_params(request.form)
            logger.info(f"Video position params extracted: {video_position_params}")

        logger.info(f"Processing file {file.filename}")
        logger.info(f"annotation_option={annotation_option}")
        logger.info(f"implement_video_frames={implement_video_frames}")

        # Process the uploaded file (always use Manual tab logic)
        result = process_uploaded_file(
            file, 
            annotation_option=annotation_option, 
            is_multi_tab=False,  # Always Manual tab logic
            implement_video_frames=implement_video_frames,
            video_position_params=video_position_params
        )

        if result['success']:
            return render_template('result.html', 
                                 filename=result['filename'],
                                 ppt_file=result['ppt_file'],
                                 folder_count=result['folder_count'],
                                 slide_count=result['slide_count'],
                                 video_folder_found=result['video_folder_found'])
        else:
            flash(result['message'], 'error')
            return redirect(url_for('index'))

    except Exception as e:
        logger.error(f"Error in upload_file: {str(e)}")
        flash(f'An error occurred: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    """Download generated files."""
    try:
        from flask import send_file
        output_folder = 'outputs'
        file_path = os.path.join(output_folder, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            flash('File not found', 'error')
            return redirect(url_for('index'))
    except Exception as e:
        logger.error(f"Error downloading file: {str(e)}")
        flash('Error downloading file', 'error')
        return redirect(url_for('index'))

# Additional routes or functions can be added below

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)