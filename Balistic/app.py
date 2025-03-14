from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import os
import uuid
import datetime
from werkzeug.utils import secure_filename
from dati_to_gensoft import transform_dati_imp_to_gensoft
import shutil

app = Flask(__name__)
app.secret_key = 'nike-transformation-tool-secret-key'

# Add template function for current year
@app.context_processor
def inject_now():
    return {'now': datetime.datetime.now}

# Configure upload and download folders
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload size

# Track transformation history
transformation_history = []

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if a file was uploaded
        if 'file' not in request.files:
            flash('No file selected', 'error')
            return redirect(request.url)
        
        file = request.files['file']
        
        # Check if user submitted without selecting a file
        if file.filename == '':
            flash('No file selected', 'error')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            # Generate a unique filename
            original_filename = secure_filename(file.filename)
            unique_id = str(uuid.uuid4())
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_{original_filename}")
            
            # Generate output filename
            output_filename = f"gensoft_{original_filename}"
            output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
            
            # Save the uploaded file
            file.save(input_path)
            
            try:
                # Transform the file
                result = transform_dati_imp_to_gensoft(input_path, output_path)
                
                # Record successful transformation
                timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                transformation_history.insert(0, {
                    'timestamp': timestamp,
                    'original_filename': original_filename,
                    'output_filename': output_filename,
                    'output_path': output_path,
                    'status': 'success',
                    'rows': len(result)
                })
                
                # Keep only the last 10 transformations in history
                if len(transformation_history) > 10:
                    transformation_history.pop()
                
                flash('File transformed successfully!', 'success')
                return redirect(url_for('download', filename=output_filename))
            
            except Exception as e:
                # Record failed transformation
                timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                transformation_history.insert(0, {
                    'timestamp': timestamp,
                    'original_filename': original_filename,
                    'status': 'error',
                    'error_message': str(e)
                })
                
                flash(f'Error during transformation: {str(e)}', 'error')
                return redirect(url_for('index'))
        else:
            flash('Only Excel files (.xlsx, .xls) are allowed', 'error')
            return redirect(request.url)
    
    return render_template('index.html', history=transformation_history)

@app.route('/download/<filename>')
def download(filename):
    """Download a generated file"""
    return send_file(os.path.join(app.config['DOWNLOAD_FOLDER'], filename),
                     as_attachment=True)

# Clean up old files periodically (could be done with a background task)
@app.route('/cleanup', methods=['POST'])
def cleanup():
    """Admin route to clean up old files"""
    # Simple implementation - just for demonstration
    for folder in [UPLOAD_FOLDER, DOWNLOAD_FOLDER]:
        for filename in os.listdir(folder):
            filepath = os.path.join(folder, filename)
            try:
                if os.path.isfile(filepath) and (datetime.datetime.now().timestamp() - os.path.getmtime(filepath)) > 86400:  # older than 1 day
                    os.remove(filepath)
            except Exception as e:
                print(f"Error cleaning up file {filepath}: {e}")
    
    flash('Cleanup completed successfully', 'success')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)