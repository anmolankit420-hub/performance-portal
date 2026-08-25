from flask import Flask, render_template, request, redirect, url_for, session, jsonify, flash
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Session secure rakhne ke liye

# Excel file path (Aapke project folder ke hisab se)
EXCEL_FILE = 'asr_tracker.xlsx'

def load_data():
    if os.path.exists(EXCEL_FILE):
        try:
            # Pandas se excel file read karein
            df = pd.read_excel(EXCEL_FILE)
            # Extra spaces hataane ke liye column names se
            df.columns = df.columns.str.strip()
            df = df.fillna('')
            return df.to_dict(orient='records')
        except Exception as e:
            print(f"Error reading Excel: {e}")
            return []
    return []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/admin_panel')
def admin_panel():
    # Check karein ki admin logged in hai ya nahi
    if 'admin_logged_in' not in session:
        return redirect(url_for('admin_login'))
    
    search_query = request.args.get('search', '').strip().lower()
    agents = load_data()
    
    # File exist karti hai ya nahi check karein
    file_exists = os.path.exists(EXCEL_FILE)
    file_name = EXCEL_FILE if file_exists else ''
    
    # Search filter apply karein
    if search_query:
        agents = [a for a in agents if search_query in str(a.get('CZ ID', '')).lower() or search_query in str(a.get('Names', '')).lower()]
    
    # Active aur Inactive count nikalna
    total_active = sum(1 for a in agents if str(a.get('Status', '')).strip().lower() != 'inactive')
    total_inactive = sum(1 for a in agents if str(a.get('Status', '')).strip().lower() == 'inactive')
    
    return render_template('admin_panel.html', 
                           agents=agents, 
                           total_active=total_active, 
                           total_inactive=total_inactive,
                           search=search_query,
                           file_exists=file_exists,
                           file_name=file_name)

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        # Apna admin username/password yahan set kar sakte hain
        if username == 'admin' and password == 'admin123':
            session['admin_logged_in'] = True
            return redirect(url_for('admin_panel'))
        else:
            flash('Invalid username or password!', 'error')
    return render_template('admin_login.html')

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_logged_in', None)
    return redirect(url_for('admin_login'))

@app.route('/admin/upload_excel', methods=['POST'])
def upload_excel():
    if 'admin_logged_in' not in session:
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    if 'excel_file' not in request.files:
        return jsonify({'success': False, 'error': 'No file part'})
    
    file = request.files['excel_file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No selected file'})
    
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
        try:
            file.save(EXCEL_FILE)
            return jsonify({'success': True})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})
            
    return jsonify({'success': False, 'error': 'Invalid file format. Only .xlsx or .xls allowed.'})

@app.route('/admin/delete_excel', methods=['POST'])
def delete_excel():
    if 'admin_logged_in' not in session:
        return redirect(url_for('admin_login'))
    
    if os.path.exists(EXCEL_FILE):
        try:
            os.remove(EXCEL_FILE)
            flash('Excel file deleted successfully!', 'success')
        except Exception as e:
            flash(f'Error deleting file: {e}', 'error')
    else:
        flash('No file found to delete.', 'error')
        
    return redirect(url_for('admin_panel'))

@app.route('/admin/download_excel')
def download_excel():
    if 'admin_logged_in' not in session:
        return redirect(url_for('admin_login'))
    
    if os.path.exists(EXCEL_FILE):
        from flask import send_file
        return send_file(EXCEL_FILE, as_attachment=True)
    else:
        flash('No Excel file available for download.', 'error')
        return redirect(url_for('admin_panel'))

# Add agent, toggle status, delete agent routes aapke pehle wale hi rahenge ya zaroorat ho toh bata dena.

if __name__ == '__main__':
    app.run(debug=True)
