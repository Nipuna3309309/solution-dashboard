from flask import Flask, render_template, request, redirect, url_for, session, send_file, jsonify, send_from_directory
from functools import wraps
import os

app = Flask(__name__, static_folder='.', static_url_path='')
app.secret_key = 'solution-dashboard-secret-key-2024'

# Authentication credentials
USERS = {
    'nipuna': 'Abey@3309309'
}

EXCEL_FILE = 'Solution List.xlsx'

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')

        if username in USERS and USERS[username] == password:
            session['logged_in'] = True
            session['username'] = username
            return redirect(url_for('dashboard'))
        else:
            error = 'Invalid credentials. Please try again.'

    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def dashboard():
    return send_file('index.html')

@app.route('/download-excel')
@login_required
def download_excel():
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True, download_name=EXCEL_FILE)
    return jsonify({'error': 'File not found'}), 404

@app.route('/upload-excel', methods=['POST'])
@login_required
def upload_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400

    if file and file.filename.endswith('.xlsx'):
        file.save(EXCEL_FILE)
        return jsonify({'success': True, 'message': 'File uploaded successfully!'})

    return jsonify({'error': 'Invalid file type. Please upload an Excel file (.xlsx)'}), 400

@app.route('/check-auth')
def check_auth():
    if 'logged_in' in session:
        return jsonify({'authenticated': True, 'username': session.get('username')})
    return jsonify({'authenticated': False})

# Serve static files
@app.route('/<path:filename>')
@login_required
def serve_static(filename):
    return send_from_directory('.', filename)

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
