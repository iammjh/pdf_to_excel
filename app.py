"""
app.py – Simple web interface for pdf_to_excel.
Upload one or more SDS PDFs, download the extracted Excel file.
"""

import os
import io
import tempfile

from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

# Import extraction logic from the same folder
from pdf_to_excel import extract_from_pdf, autofit_worksheet

import pandas as pd

app = Flask(__name__)
app.secret_key = 'sds-extractor-secret'

ALLOWED_EXT = {'pdf'}


def allowed(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/extract', methods=['POST'])
def extract():
    files = request.files.getlist('pdfs')
    if not files or all(f.filename == '' for f in files):
        flash('Please select at least one PDF file.')
        return redirect(url_for('index'))

    rows = []

    with tempfile.TemporaryDirectory() as tmpdir:
        for f in files:
            if not allowed(f.filename):
                flash(f'"{f.filename}" is not a PDF — skipped.')
                continue

            # Save upload to temp file
            safe_name = secure_filename(f.filename) or 'uploaded.pdf'
            tmp_path = os.path.join(tmpdir, safe_name)
            f.save(tmp_path)

            try:
                result = extract_from_pdf(tmp_path)
                product_name = result.get('product_name') or safe_name

                items = result.get('items')
                if items is None:
                    items = [
                        {'chem_name': '', 'cas': cas}
                        for cas in result.get('cas_list', [])
                    ]

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
            except Exception as exc:
                flash(f'Error processing "{f.filename}": {exc}')
                rows.append({
                    'Product Name':  f.filename,
                    'Chemical Name': '',
                    'CAS Number':    f'ERROR: {exc}',
                })

    if not rows:
        flash('No data could be extracted.')
        return redirect(url_for('index'))

    df = pd.DataFrame(rows, columns=['Product Name', 'Chemical Name', 'CAS Number'])

    # Write Excel to memory buffer and send as download
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Chemicals')
        autofit_worksheet(writer.sheets['Chemicals'])
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name='chemicals_output.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


if __name__ == '__main__':
    app.run(debug=True, port=5000)
