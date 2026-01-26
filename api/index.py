from flask import Flask, render_template_string, request, redirect, url_for, session, send_file, jsonify, Response
from functools import wraps
import os
import json

app = Flask(__name__)
app.secret_key = 'solution-dashboard-secret-key-2024'

# Authentication credentials
USERS = {
    'nipuna': 'Abey@3309309'
}

# Get the base directory
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_FILE = os.path.join(BASE_DIR, 'Solution List.xlsx')

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            return redirect('/login')
        return f(*args, **kwargs)
    return decorated_function

# Login page HTML
LOGIN_HTML = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login - Solution Dashboard</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%); min-height: 100vh; display: flex; align-items: center; justify-content: center; }
        .login-container { background: #fff; border-radius: 16px; box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3); overflow: hidden; width: 400px; max-width: 90%; }
        .login-header { background: linear-gradient(135deg, #252423 0%, #3d3d3d 100%); padding: 30px; text-align: center; }
        .login-header h1 { color: #F2C811; font-size: 24px; margin-bottom: 8px; }
        .login-header p { color: #aaa; font-size: 14px; }
        .logo-icon { width: 60px; height: 60px; background: #F2C811; border-radius: 12px; display: flex; align-items: center; justify-content: center; margin: 0 auto 15px; }
        .logo-icon svg { width: 35px; height: 35px; fill: #252423; }
        .login-form { padding: 40px 30px; }
        .form-group { margin-bottom: 20px; }
        .form-group label { display: block; margin-bottom: 8px; color: #333; font-weight: 600; font-size: 14px; }
        .form-group input { width: 100%; padding: 14px 16px; border: 2px solid #e0e0e0; border-radius: 8px; font-size: 15px; transition: all 0.3s ease; }
        .form-group input:focus { outline: none; border-color: #F2C811; box-shadow: 0 0 0 3px rgba(242, 200, 17, 0.2); }
        .error-message { background: #fee2e2; color: #dc2626; padding: 12px 16px; border-radius: 8px; margin-bottom: 20px; font-size: 14px; }
        .login-btn { width: 100%; padding: 14px; background: linear-gradient(135deg, #F2C811 0%, #e6b800 100%); color: #252423; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; transition: all 0.3s ease; }
        .login-btn:hover { transform: translateY(-2px); box-shadow: 0 5px 20px rgba(242, 200, 17, 0.4); }
        .footer-text { text-align: center; margin-top: 20px; color: #888; font-size: 12px; }
    </style>
</head>
<body>
    <div class="login-container">
        <div class="login-header">
            <div class="logo-icon"><svg viewBox="0 0 24 24"><path d="M3 13h2v-2H3v2zm0 4h2v-2H3v2zm0-8h2V7H3v2zm4 4h14v-2H7v2zm0 4h14v-2H7v2zM7 7v2h14V7H7z"/></svg></div>
            <h1>Solution Dashboard</h1>
            <p>Sign in to access your dashboard</p>
        </div>
        <form class="login-form" method="POST" action="/login">
            {% if error %}<div class="error-message">{{ error }}</div>{% endif %}
            <div class="form-group"><label for="username">Username</label><input type="text" id="username" name="username" placeholder="Enter your username" required></div>
            <div class="form-group"><label for="password">Password</label><input type="password" id="password" name="password" placeholder="Enter your password" required></div>
            <button type="submit" class="login-btn">Sign In</button>
            <p class="footer-text">Secure access to your analytics dashboard</p>
        </form>
    </div>
</body>
</html>'''

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        if username in USERS and USERS[username] == password:
            session['logged_in'] = True
            session['username'] = username
            return redirect('/')
        else:
            error = 'Invalid credentials. Please try again.'
    return render_template_string(LOGIN_HTML, error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')

@app.route('/')
@login_required
def dashboard():
    index_path = os.path.join(BASE_DIR, 'index.html')
    with open(index_path, 'r', encoding='utf-8') as f:
        return f.read()

@app.route('/download-excel')
@login_required
def download_excel():
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True, download_name='Solution List.xlsx')
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

@app.route('/<path:filename>')
@login_required
def serve_static(filename):
    file_path = os.path.join(BASE_DIR, filename)
    if os.path.exists(file_path):
        # Determine content type
        content_type = 'text/plain'
        if filename.endswith('.html'):
            content_type = 'text/html'
        elif filename.endswith('.css'):
            content_type = 'text/css'
        elif filename.endswith('.js'):
            content_type = 'application/javascript'
        elif filename.endswith('.xlsx'):
            content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

        with open(file_path, 'rb') as f:
            return Response(f.read(), mimetype=content_type)
    return 'Not found', 404

# Vercel handler
app = app

# Local development server
if __name__ == '__main__':
    app.run(debug=True, port=5000)
