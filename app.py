"""
Flask Web Application for Excel Cleaning
Allows users to upload raw Excel files and download cleaned versions
"""

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import os
import tempfile
import time
from pathlib import Path
from clean_excel import clean_excel_file

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_FILE_SIZE = 16 * 1024 * 1024  # 16MB

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Render main page"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """
    Handle file upload and processing
    
    Form fields (all optional):
    - file: Excel file to clean
    - bed_color: Color code for Bed fabric (optional - auto-detected from file)
    - liv_color: Color code for Liv fabric (optional - auto-detected from file)
    - deduction_cell: Cell reference for D value (default: I6)
    - fabric_colors: JSON string of additional fabric colors (optional)
    """
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Please upload .xlsx or .xls files'}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        timestamp = str(int(time.time()))
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp}_{filename}")
        file.save(input_path)
        
        # Get optional configuration from form
        bed_color = request.form.get('bed_color', '').strip() or None
        liv_color = request.form.get('liv_color', '').strip() or None
        deduction_cell = request.form.get('deduction_cell', 'I6').strip() or 'I6'
        
        # Parse additional fabric colors (JSON format)
        fabric_colors = None
        fabric_colors_json = request.form.get('fabric_colors', '').strip()
        if fabric_colors_json:
            try:
                import json
                fabric_colors = json.loads(fabric_colors_json)
            except json.JSONDecodeError:
                # Silently ignore invalid JSON
                pass
        
        # Generate output path
        output_filename = f"{Path(filename).stem}-cleaned.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp}_{output_filename}")
        
        # Process the file
        clean_excel_file(
            input_file=input_path,
            output_file=output_path,
            bed_color=bed_color,
            liv_color=liv_color,
            fabric_colors=fabric_colors,
            deduction_cell=deduction_cell
        )
        
        # Clean up input file
        try:
            os.remove(input_path)
        except Exception:
            pass  # Ignore cleanup errors
        
        # Return download info
        return jsonify({
            'success': True,
            'download_id': f"{timestamp}_{output_filename}",
            'filename': output_filename
        })
    
    except Exception as e:
        # Clean up files on error
        try:
            if 'input_path' in locals():
                os.remove(input_path)
            if 'output_path' in locals() and os.path.exists(output_path):
                os.remove(output_path)
        except Exception:
            pass
        
        return jsonify({
            'error': f'Error processing file: {str(e)}'
        }), 500


@app.route('/download/<download_id>')
def download_file(download_id):
    """
    Download cleaned file
    """
    try:
        # Secure the filename
        safe_id = secure_filename(download_id)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_id)
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found or expired'}), 404
        
        # Send file and schedule cleanup
        response = send_file(
            file_path,
            as_attachment=True,
            download_name=download_id.split('_', 1)[1] if '_' in download_id else download_id
        )
        
        # Clean up file after a delay (let download complete)
        @response.call_on_close
        def cleanup():
            try:
                # Small delay to ensure file transfer completes
                time.sleep(2)
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception:
                pass  # Ignore cleanup errors
        
        return response
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({'status': 'healthy'})


if __name__ == '__main__':
    # Run in debug mode for local development
    # For production, use gunicorn: gunicorn app:app
    app.run(debug=True, host='0.0.0.0', port=5000)
