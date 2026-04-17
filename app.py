"""
app.py – Simple web interface for pdf_to_excel.
Upload one or more SDS PDFs, download the extracted Excel file.
"""

import os
import io
import tempfile
import gc
import sys

from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

# Import extraction logic from the same folder
from pdf_to_excel import extract_from_pdf, autofit_worksheet

import pandas as pd

app = Flask(__name__)
app.secret_key = 'sds-extractor-secret'

# Extend timeout for Render
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max upload

ALLOWED_EXT = {'pdf'}


def allowed(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/extract', methods=['POST'])
def extract():
    try:
        files = request.files.getlist('pdfs')
        if not files or all(f.filename == '' for f in files):
            flash('Please select at least one PDF file.')
            return redirect(url_for('index'))

        valid_files = [f for f in files if allowed(f.filename)]
        
        if not valid_files:
            flash('No valid PDF files to process.')
            return redirect(url_for('index'))

        rows = []
        processed_count = 0

        with tempfile.TemporaryDirectory() as tmpdir:
            for idx, f in enumerate(valid_files):
                try:
                    safe_name = secure_filename(f.filename) or 'uploaded.pdf'
                    tmp_path = os.path.join(tmpdir, safe_name)
                    
                    # Save file
                    f.save(tmp_path)

                    # Extract data
                    result = extract_from_pdf(tmp_path)
                    product_name = result.get('product_name') or safe_name
                    items = result.get('items', [])

                    if items:
                        for item in items:
                            rows.append({
                                'Product Name':  product_name,
                                'Chemical Name': item['chem_name'],
                                'CAS Number':    item['cas'],
                            })
                    else:
                        rows.append({
                            'Product Name':  product_name,
                            'Chemical Name': '',
                            'CAS Number':    'N/A',
                        })
                    
                    processed_count += 1
                    
                    # Clean up temp file immediately
                    try:
                        os.remove(tmp_path)
                    except:
                        pass
                    
                    # Force garbage collection every 3 files
                    if (idx + 1) % 3 == 0:
                        gc.collect()
                        
                except Exception as exc:
                    flash(f'Error processing "{f.filename}": {str(exc)[:80]}')
                    rows.append({
                        'Product Name':  f.filename,
                        'Chemical Name': '',
                        'CAS Number':    f'ERROR',
                    })

        if not rows:
            flash('No data could be extracted from any files.')
            return redirect(url_for('index'))

        # Create Excel
        df = pd.DataFrame(rows, columns=['Product Name', 'Chemical Name', 'CAS Number'])

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Chemicals')
            autofit_worksheet(writer.sheets['Chemicals'])
        buf.seek(0)

        flash(f'Successfully processed {processed_count} of {len(valid_files)} files.')
        
        return send_file(
            buf,
            as_attachment=True,
            download_name='chemicals_output.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        
    except Exception as exc:
        print(f'FATAL ERROR: {exc}', file=sys.stderr)
        flash(f'Server error: {str(exc)[:100]}')
        return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True, port=5000)
