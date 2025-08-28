from flask import Flask, render_template, request, send_file, jsonify, session
import pandas as pd
import os
from werkzeug.utils import secure_filename
from datetime import datetime
import tempfile
from excel_processor import ExcelProcessor
import uuid

app = Flask(__name__)
app.secret_key = 'excel_processing_secret_key_2024'

# Konfigurasi upload
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'Tidak ada file yang dipilih'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Tidak ada file yang dipilih'}), 400
    
    if file and allowed_file(file.filename):
        # Generate unique session ID
        session_id = str(uuid.uuid4())
        session['session_id'] = session_id
        
        # Secure filename
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_filename = f"{timestamp}_{session_id}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
        
        # Save file
        file.save(filepath)
        
        print(f"📁 File saved: {filepath}")
        print(f"📊 File size: {os.path.getsize(filepath)} bytes")
        
        try:
            # Process Excel file
            processor = ExcelProcessor()
            print(f"🔍 Starting Excel processing...")
            
            # Check if file is readable
            if not os.path.exists(filepath):
                raise Exception("File tidak ditemukan setelah upload")
            
            file_size = os.path.getsize(filepath)
            if file_size == 0:
                raise Exception("File kosong (0 bytes)")
            
            print(f"📊 File size: {file_size} bytes")
            
            # Try to open the file to check if it's a valid Excel file
            try:
                import pandas as pd
                test_df = pd.read_excel(filepath, nrows=1)
                print(f"✅ File is readable Excel file with {len(test_df.columns)} columns")
            except Exception as excel_error:
                raise Exception(f"File bukan file Excel yang valid: {str(excel_error)}")
            
            preview_data = processor.preview_excel(filepath)
            
            # Debug: Print preview data structure
            print(f"🔍 Preview data keys: {list(preview_data.keys()) if preview_data else 'None'}")
            if preview_data and 'sheets' in preview_data:
                print(f"📊 Sheets found: {list(preview_data['sheets'].keys())}")
                for sheet_name, sheet_data in preview_data['sheets'].items():
                    print(f"  📋 Sheet '{sheet_name}': {len(sheet_data.get('detected_fields', {}))} fields detected")
            
            # Validate preview data
            if not preview_data:
                raise Exception("Data preview kosong")
            if 'sheets' not in preview_data:
                raise Exception("Struktur data tidak valid - tidak ada sheets yang terdeteksi")
            if not preview_data['sheets']:
                raise Exception("Tidak ada sheet yang dapat diproses")
            
            # Additional validation: check if any fields were detected
            total_fields = sum(len(sheet.get('detected_fields', {})) for sheet in preview_data['sheets'].values())
            if total_fields == 0:
                print("⚠️ Warning: No fields detected in any sheet")
                # Don't fail here, just warn - some files might not have recognizable fields
            
            # Store filepath in session
            session['uploaded_file'] = filepath
            
            return jsonify({
                'success': True,
                'message': 'File berhasil diupload dan diproses',
                'preview': preview_data,
                'filename': filename
            })
            
        except Exception as e:
            # Log the full error for debugging
            import traceback
            print(f"❌ Error during Excel processing:")
            print(f"   Error: {str(e)}")
            print(f"   Traceback:")
            traceback.print_exc()
            
            # Clean up file if processing fails
            try:
                if os.path.exists(filepath):
                    # Close any open file handles first
                    import gc
                    gc.collect()
                    os.remove(filepath)
            except PermissionError:
                # If file is still in use, just log it
                print(f"Warning: Could not delete file {filepath} - file may still be in use")
            except Exception as cleanup_error:
                print(f"Warning: Error during cleanup: {cleanup_error}")
            
            error_msg = str(e)
            # Clean error message untuk JSON
            error_msg = error_msg.replace('\x00', '').replace('\n', ' ').replace('\r', '')
            return jsonify({'error': f'Error memproses file: {error_msg}'}), 500
    
    return jsonify({'error': 'Format file tidak didukung. Gunakan file Excel (.xlsx atau .xls)'}), 400

@app.route('/process', methods=['POST'])
def process_excel():
    if 'uploaded_file' not in session:
        return jsonify({'error': 'Tidak ada file yang diupload'}), 400
    
    filepath = session['uploaded_file']
    
    if not os.path.exists(filepath):
        return jsonify({'error': 'File tidak ditemukan'}), 400
    
    try:
        # Get processing options from request
        data = request.get_json()
        options = data.get('options', {})
        
        # Process Excel file
        processor = ExcelProcessor()
        output_filepath = processor.process_excel(filepath, options)
        
        # Store output filepath in session
        session['output_file'] = output_filepath
        
        return jsonify({
            'success': True,
            'message': 'File berhasil diproses',
            'output_filename': os.path.basename(output_filepath)
        })
        
    except Exception as e:
        return jsonify({'error': f'Error memproses file: {str(e)}'}), 500

@app.route('/download')
def download_file():
    if 'output_file' not in session:
        return jsonify({'error': 'Tidak ada file output yang tersedia'}), 400
    
    output_filepath = session['output_file']
    
    if not os.path.exists(output_filepath):
        return jsonify({'error': 'File output tidak ditemukan'}), 400
    
    try:
        return send_file(
            output_filepath,
            as_attachment=True,
            download_name=os.path.basename(output_filepath),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': f'Error download file: {str(e)}'}), 500

@app.route('/cleanup', methods=['POST'])
def cleanup_files():
    """Clean up uploaded and processed files"""
    try:
        # Clean up uploaded file
        if 'uploaded_file' in session and os.path.exists(session['uploaded_file']):
            try:
                os.remove(session['uploaded_file'])
                del session['uploaded_file']
            except PermissionError:
                print(f"Warning: Could not delete uploaded file {session['uploaded_file']}")
                del session['uploaded_file']
        
        # Clean up output file
        if 'output_file' in session and os.path.exists(session['output_file']):
            try:
                os.remove(session['output_file'])
                del session['output_file']
            except PermissionError:
                print(f"Warning: Could not delete output file {session['output_file']}")
                del session['output_file']
        
        return jsonify({'success': True, 'message': 'File berhasil dibersihkan'})
    except Exception as e:
        return jsonify({'error': f'Error membersihkan file: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
