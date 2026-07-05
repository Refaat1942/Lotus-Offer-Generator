import os
import secrets
import uuid
from datetime import datetime
from functools import wraps

from flask import (
    Flask, jsonify, redirect, render_template, request,
    send_file, session, url_for
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename
import io

from offer_engine import (
    APP_VERSION, BRANCHES, DEFAULT_OFFERS,
    create_template_bytes, export_to_excel, get_companies_from_df,
    process_mix_match, read_sales_file
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'static', 'uploads')
LOGO_PATH = os.path.join(UPLOAD_DIR, 'logo.png')
DATA_DIR = os.path.join(BASE_DIR, 'data')

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', secrets.token_hex(32))
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

ADMIN_USER = os.environ.get('ADMIN_USER', 'admin')
ADMIN_PASS_HASH = generate_password_hash(os.environ.get('ADMIN_PASS', 'admin'))


def _session_file_path(key):
    sid = session.get('session_id')
    if not sid:
        return None
    return os.path.join(DATA_DIR, f'{sid}_{key}')


def _ensure_session_id():
    if 'session_id' not in session:
        session['session_id'] = uuid.uuid4().hex


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Unauthorized'}), 401
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('logged_in'):
        return redirect(url_for('index'))
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        if username == ADMIN_USER and check_password_hash(ADMIN_PASS_HASH, password):
            session['logged_in'] = True
            session['username'] = username
            _ensure_session_id()
            return redirect(url_for('index'))
        error = 'Invalid username or password'
    return render_template('login.html', error=error, version=APP_VERSION)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/')
@login_required
def index():
    has_logo = os.path.exists(LOGO_PATH)
    return render_template(
        'index.html',
        version=APP_VERSION,
        branches=BRANCHES,
        default_offers=DEFAULT_OFFERS,
        has_logo=has_logo
    )


@app.route('/api/logo', methods=['GET'])
def get_logo():
    if os.path.exists(LOGO_PATH):
        return send_file(LOGO_PATH, mimetype='image/png')
    return '', 404


@app.route('/api/logo', methods=['POST'])
@login_required
def upload_logo():
    if 'logo' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    file = request.files['logo']
    if not file.filename:
        return jsonify({'error': 'No file selected'}), 400
    allowed = {'png', 'jpg', 'jpeg', 'webp', 'svg'}
    ext = file.filename.rsplit('.', 1)[-1].lower()
    if ext not in allowed:
        return jsonify({'error': 'Invalid file type. Use PNG, JPG, or WEBP.'}), 400
    file.save(LOGO_PATH)
    return jsonify({'success': True, 'message': 'Logo uploaded successfully'})


@app.route('/api/template')
@login_required
def download_template():
    data = create_template_bytes()
    return send_file(
        io.BytesIO(data),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Lotus_Sales_Template.xlsx'
    )


@app.route('/api/upload', methods=['POST'])
@login_required
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    file = request.files['file']
    if not file.filename:
        return jsonify({'error': 'No file selected'}), 400

    try:
        file_bytes = file.read()
        df = read_sales_file(file_bytes, file.filename)
        companies = get_companies_from_df(df)
        _ensure_session_id()
        upload_path = _session_file_path('upload')
        with open(upload_path, 'wb') as f:
            f.write(file_bytes)
        session['uploaded_filename'] = secure_filename(file.filename)
        return jsonify({
            'success': True,
            'companies': companies,
            'filename': file.filename,
            'row_count': len(df)
        })
    except Exception as e:
        return jsonify({'error': f'Could not read file: {str(e)}'}), 400


@app.route('/api/process', methods=['POST'])
@login_required
def process():
    upload_path = _session_file_path('upload')
    if not upload_path or not os.path.exists(upload_path):
        return jsonify({'error': 'Please upload a file first'}), 400

    data = request.get_json()
    if not data:
        return jsonify({'error': 'Invalid request'}), 400

    company_mappings = data.get('company_mappings', {})
    target_branch = data.get('branch', 'All')
    use_date_filter = data.get('filter_by_date', False)
    start_date = data.get('start_date')
    end_date = data.get('end_date')
    target_str = str(data.get('target_discount', '')).strip()

    try:
        target_discount = float(target_str) if target_str else float('inf')
    except ValueError:
        return jsonify({'error': 'Please enter a valid number for Target Discount'}), 400

    try:
        with open(upload_path, 'rb') as f:
            file_bytes = f.read()
        filename = session.get('uploaded_filename', 'data.xlsx')
        df_original = read_sales_file(file_bytes, filename)

        if use_date_filter and start_date and end_date:
            start_date = datetime.strptime(start_date, '%Y-%m-%d').date()
            end_date = datetime.strptime(end_date, '%Y-%m-%d').date()
        else:
            start_date = None
            end_date = None

        processed_data, true_unprocessed, accumulated_discount = process_mix_match(
            df_original,
            company_mappings,
            target_branch=target_branch,
            use_date_filter=use_date_filter,
            start_date=start_date,
            end_date=end_date,
            target_discount=target_discount,
        )

        total_processed = sum(len(lst) for lst in processed_data.values())
        if total_processed == 0 and not true_unprocessed:
            return jsonify({'error': 'No data generated. Check your mappings.'}), 400

        excel_bytes = export_to_excel(processed_data, true_unprocessed)
        result_path = _session_file_path('result')
        with open(result_path, 'wb') as f:
            f.write(excel_bytes)

        target_display = target_str if target_str else 'No Limit'
        return jsonify({
            'success': True,
            'version': APP_VERSION,
            'target_discount': target_display,
            'achieved_discount': f'{accumulated_discount:,.2f}',
            'unprocessed_count': len(true_unprocessed),
            'processed_count': total_processed,
            'offer_sheets': list(processed_data.keys())
        })
    except Exception as e:
        return jsonify({'error': f'Processing failed: {str(e)}'}), 500


@app.route('/api/download')
@login_required
def download_result():
    result_path = _session_file_path('result')
    if not result_path or not os.path.exists(result_path):
        return jsonify({'error': 'No processed file available'}), 404
    filename = f"Lotus_MixMatch_v{APP_VERSION}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(
        result_path,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/api/clear', methods=['POST'])
@login_required
def clear_data():
    for key in ('upload', 'result'):
        path = _session_file_path(key)
        if path and os.path.exists(path):
            os.remove(path)
    session.pop('uploaded_filename', None)
    return jsonify({'success': True})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 17800))
    app.run(host='0.0.0.0', port=port, debug=False)
