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
from clean_excel import clean_excel_file, scan_text_tags

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


@app.route('/scan', methods=['POST'])
def scan_file():
    """
    Phase 1: Upload file and scan for non-numeric tag rows.

    Returns JSON:
      { "has_text_tags": bool, "rows": [...], "scan_id": "timestamp_filename.xlsx" }

    The file is kept on disk so /upload can reuse it via scan_id.
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Please upload .xlsx or .xls files'}), 400

        filename = secure_filename(file.filename)
        timestamp = str(int(time.time()))
        scan_id = f"{timestamp}_{filename}"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], scan_id)
        file.save(input_path)

        # Run the quick scan
        scan_result = scan_text_tags(input_path)
        
        # Handle dict format vs legacy list format just in case
        if isinstance(scan_result, dict):
            text_tag_rows = scan_result.get('text_rows', [])
            missing_headers = scan_result.get('missing_headers', [])
        else:
            text_tag_rows = scan_result
            missing_headers = []

        # Sanitise values for JSON serialisation
        import math
        def _safe(v):
            try:
                if isinstance(v, float) and math.isnan(v):
                    return ''
                return v
            except Exception:
                return str(v) if v is not None else ''

        rows_json = [
            {k: _safe(v) for k, v in row.items()}
            for row in text_tag_rows
        ]

        return jsonify({
            'has_missing_headers': len(missing_headers) > 0,
            'missing_headers': missing_headers,
            'has_text_tags': len(rows_json) > 0,
            'rows': rows_json,
            'scan_id': scan_id
        })

    except Exception as e:
        return jsonify({'error': f'Error scanning file: {str(e)}'}), 500


@app.route('/upload', methods=['POST'])
def upload_file():
    """
    Phase 2: Process the file and return cleaned output.

    Accepts either:
      - A fresh file upload (no scan_id)  — backwards-compatible
      - scan_id (string) referencing a file already saved by /scan

    Optional form fields:
      - scan_id:       Reference to pre-saved file from /scan
      - tag_action:    'skip' (default) | 'keep' | 'extract'
      - bed_color, liv_color, deduction_cell, fabric_colors
    """
    try:
        scan_id = request.form.get('scan_id', '').strip()
        tag_action = request.form.get('tag_action', 'skip').strip() or 'skip'

        if scan_id:
            # Reuse file saved during /scan
            safe_id = secure_filename(scan_id)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_id)
            if not os.path.exists(input_path):
                return jsonify({'error': 'Scanned file not found or expired. Please re-upload.'}), 400
            filename = safe_id.split('_', 1)[1] if '_' in safe_id else safe_id
        else:
            # Fresh upload (no prior scan)
            if 'file' not in request.files:
                return jsonify({'error': 'No file uploaded'}), 400
            file = request.files['file']
            if file.filename == '':
                return jsonify({'error': 'No file selected'}), 400
            if not allowed_file(file.filename):
                return jsonify({'error': 'Invalid file type. Please upload .xlsx or .xls files'}), 400

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
                pass

        # Generate output path
        stem = Path(filename).stem
        timestamp_out = str(int(time.time()))
        output_filename = f"{stem}-cleaned.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp_out}_{output_filename}")

        # Process the file
        clean_excel_file(
            input_file=input_path,
            output_file=output_path,
            bed_color=bed_color,
            liv_color=liv_color,
            fabric_colors=fabric_colors,
            deduction_cell=deduction_cell,
            tag_action=tag_action
        )

        # Clean up input file
        try:
            os.remove(input_path)
        except Exception:
            pass

        return jsonify({
            'success': True,
            'download_id': f"{timestamp_out}_{output_filename}",
            'filename': output_filename
        })

    except Exception as e:
        try:
            if 'input_path' in locals():
                os.remove(input_path)
            if 'output_path' in locals() and os.path.exists(output_path):
                os.remove(output_path)
        except Exception:
            pass

        return jsonify({'error': f'Error processing file: {str(e)}'}), 500


@app.route('/download/<download_id>')
def download_file(download_id):
    """Download cleaned file"""
    try:
        safe_id = secure_filename(download_id)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_id)

        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found or expired'}), 404

        response = send_file(
            file_path,
            as_attachment=True,
            download_name=download_id.split('_', 1)[1] if '_' in download_id else download_id
        )

        @response.call_on_close
        def cleanup():
            try:
                time.sleep(2)
                if os.path.exists(file_path):
                    os.remove(file_path)
            except Exception:
                pass

        return response

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({'status': 'healthy'})


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
