import os
import tempfile
import logging
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix

# Configure logging for serverless environment
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Create Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key-change-in-production")
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

# Configuration
UPLOAD_FOLDER = '/tmp/uploads'
OUTPUT_FOLDER = '/tmp/outputs'
MAX_CONTENT_LENGTH = 200 * 1024 * 1024  # 200MB max file size
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'zip'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def allowed_file(filename):
    """Check if file has allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Main page with upload form."""
    try:
        logger.info("Accessing main page")
        return render_template('index.html')
    except Exception as e:
        logger.error(f"Error in index route: {str(e)}")
        return f"Error: {str(e)}", 500


@app.route('/health')
def health_check():
    """Health check endpoint for debugging."""
    try:
        return jsonify({
            'status': 'healthy',
            'message': 'Flask app is running in serverless environment',
            'upload_folder': UPLOAD_FOLDER,
            'output_folder': OUTPUT_FOLDER,
            'temp_dir': tempfile.gettempdir()
        })
    except Exception as e:
        logger.error(f"Error in health check: {str(e)}")
        return jsonify({'error': str(e)}), 500


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload - simplified version."""
    try:
        logger.info("Upload endpoint called")
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file selected'}), 400

        file = request.files['file']
        
        if not file or file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type'}), 400

        # Save file temporarily
        filename = secure_filename(file.filename)
        temp_path = os.path.join(tempfile.gettempdir(), filename)
        file.save(temp_path)
        
        # For now, just return success without processing
        file_size = os.path.getsize(temp_path)
        
        # Clean up
        os.remove(temp_path)
        
        return jsonify({
            'success': True,
            'message': 'File uploaded successfully',
            'filename': filename,
            'size': file_size
        })
        
    except Exception as e:
        logger.error(f"Error in upload: {str(e)}")
        return jsonify({'error': f'Upload error: {str(e)}'}), 500


@app.route('/manual-upload', methods=['POST'])
def manual_upload():
    """Handle manual upload - simplified version."""
    return upload_file()  # Same logic for now


@app.route('/download/<filename>')
def download_file(filename):
    """Download generated files - placeholder."""
    return jsonify({'error': 'Download not implemented in simplified version'}), 501


@app.route('/convert-to-pdf/<filename>', methods=['GET', 'POST'])
def convert_to_pdf(filename):
    """Convert PowerPoint to PDF - placeholder."""
    return jsonify({'error': 'PDF conversion not implemented in simplified version'}), 501


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)