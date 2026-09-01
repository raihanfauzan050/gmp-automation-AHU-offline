"""
GMP Automation System - Main Web Application
Flask-based web interface for processing environmental measurement PDFs.
"""

import os
import json
import uuid
import traceback
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from werkzeug.utils import secure_filename
from ahu_utils import ahu_sort_key, default_ahu_for_test, extract_ahu_number
from config import UPLOAD_FOLDER, OUTPUT_FOLDER, get_semester_label, TEST_TYPES
from deepseek_ocr.engine import EXTRACTORS as DEEPSEEK_EXTRACTORS
from excel_generator import GENERATORS

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}

ERROR_MESSAGES = {
    'ko': {
        'invalid_test': '잘못된 측정 종류가 선택되었습니다.',
        'no_files': '업로드된 PDF 파일이 없습니다.',
        'no_valid_files': '유효한 PDF 파일을 찾을 수 없습니다.',
        'extract_failed': '모든 PDF에서 데이터를 추출하지 못했습니다.',
        'file_not_found': '파일을 찾을 수 없습니다.',
        'server_error': '서버 오류:',
        'processing_error': '처리 오류:',
        'empty_data': '측정표에서 데이터를 찾을 수 없습니다.',
    },
    'en': {
        'invalid_test': 'Invalid test type selected.',
        'no_files': 'No PDF files uploaded.',
        'no_valid_files': 'No valid PDF files found.',
        'extract_failed': 'Failed to extract data from all PDFs.',
        'file_not_found': 'File not found.',
        'server_error': 'Server error:',
        'processing_error': 'Error processing:',
        'empty_data': 'No measurement data was found in the table.',
    },
}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Redirect the root URL to the offline OCR workflow."""
    return redirect(url_for('offline'))


@app.route('/offline')
def offline():
    """Offline OCR workflow using the configured OCR endpoint."""
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    """Process uploaded PDFs and generate Excel files."""
    try:
        test_type = request.form.get('test_type')
        language = request.form.get('language', 'ko').strip()
        messages = ERROR_MESSAGES.get(language, ERROR_MESSAGES['ko'])
        if not test_type or test_type not in DEEPSEEK_EXTRACTORS:
            return jsonify({'error': messages['invalid_test']}), 400

        files = request.files.getlist('pdf_files')
        if not files or all(f.filename == '' for f in files):
            return jsonify({'error': messages['no_files']}), 400

        # Save uploaded files
        saved_paths = []
        for f in files:
            if f and allowed_file(f.filename):
                filename = secure_filename(f.filename)
                unique_name = f"{uuid.uuid4().hex}_{filename}"
                filepath = os.path.join(UPLOAD_FOLDER, unique_name)
                f.save(filepath)
                saved_paths.append(filepath)

        if not saved_paths:
            return jsonify({'error': messages['no_valid_files']}), 400

        # Extract data from each PDF
        extractor = DEEPSEEK_EXTRACTORS[test_type]
        all_ahu_data = {}
        errors = []

        for pdf_path in saved_paths:
            try:
                data = extractor(pdf_path)
                data_key = {
                    'airborne_particle': 'rooms',
                    'air_velocity': 'machines',
                    'air_change_rate': 'rooms',
                    'hepa_filter': 'items',
                    'airflow_pattern': 'items',
                }[test_type]
                if not data.get(data_key):
                    raise ValueError(messages['empty_data'])
                ahu_num = extract_ahu_number(
                    data.get('ahu'),
                    pdf_path,
                    default=default_ahu_for_test(test_type),
                )
                date_str = data.get('date')
                if not date_str and test_type == 'airflow_pattern':
                    date_str = data['items'][0].get('date')
                date_str = date_str or '2025.08.01'
                semester_label = get_semester_label(date_str)

                # Organize data by AHU
                if ahu_num not in all_ahu_data:
                    all_ahu_data[ahu_num] = []

                # Build semester data based on test type
                sem_entry = {'semester': semester_label, 'date': date_str}

                if test_type == 'airborne_particle':
                    sem_entry['rooms'] = data.get('rooms', [])
                elif test_type == 'air_velocity':
                    sem_entry['machines'] = data.get('machines', [])
                elif test_type == 'air_change_rate':
                    sem_entry['rooms'] = data.get('rooms', [])
                elif test_type == 'hepa_filter':
                    sem_entry['items'] = data.get('items', [])
                elif test_type == 'airflow_pattern':
                    sem_entry['items'] = data.get('items', [])

                all_ahu_data[ahu_num].append(sem_entry)

            except Exception as e:
                errors.append(f"{messages['processing_error']} {os.path.basename(pdf_path)}: {str(e)}")

        if not all_ahu_data:
            error_msg = messages['extract_failed']
            if errors:
                error_msg += "\n" + "\n".join(errors)
            return jsonify({'error': error_msg}), 400

        # Generate Excel
        generator = GENERATORS[test_type]
        test_config = TEST_TYPES[test_type]
        output_filename = test_config['excel_filename']
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        generator(all_ahu_data, output_path)

        # Clean up uploaded files
        for p in saved_paths:
            try:
                os.remove(p)
            except:
                pass

        result = {
            'success': True,
            'filename': output_filename,
            'download_url': f'/download/{output_filename}',
            'ahu_count': len(all_ahu_data),
            'ahu_list': sorted(all_ahu_data.keys(), key=ahu_sort_key),
        }

        if errors:
            result['warnings'] = errors

        return jsonify(result)

    except Exception as e:
        traceback.print_exc()
        language = request.form.get('language', 'ko').strip()
        messages = ERROR_MESSAGES.get(language, ERROR_MESSAGES['ko'])
        return jsonify({'error': f"{messages['server_error']} {str(e)}"}), 500


@app.route('/download/<filename>')
def download(filename):
    """Download generated Excel file."""
    filepath = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=filename)
    return jsonify({'error': ERROR_MESSAGES['ko']['file_not_found']}), 404


if __name__ == '__main__':
    print("=" * 60)
    print("  GMP Automation System")
    print("  Open your browser and go to: http://localhost:5002/offline")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5002, debug=False)
