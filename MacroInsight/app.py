from flask import Flask, render_template, request, redirect, url_for, flash
import os
from werkzeug.utils import secure_filename
from vba_analysis.extract import extract_vba_code, analyze_vba_code

app = Flask(__name__)
app.secret_key = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = 'uploads/'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    if file and file.filename.endswith('.xlsm'):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        vba_code = extract_vba_code(file_path)
        if vba_code:
            analysis_results = analyze_vba_code(vba_code)
            return render_template('results.html', vba_code=vba_code, analysis_results=analysis_results)
        else:
            flash('Failed to extract VBA code from the file.')
            return redirect(request.url)
    else:
        flash('Invalid file format. Please upload an XLSM file.')
        return redirect(request.url)

if __name__ == '__main__':
    app.run(debug=True)
