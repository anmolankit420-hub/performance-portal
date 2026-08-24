import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = 'asr_secret_key_12345'
EXCEL_FILE = 'asr_tracker.xlsx'

def load_data():
    desired_columns = [
        'CZ ID', 'Names', 'Shift', 'TL', 'QA', 'PIP', 'June', 'July', 'MTD Aug', 
        'D-2', 'D-1', 'D-Day', 'D-Day SOB POC%', 'Mandays', 'CPA', 
        'Target-Booking%', 'Booking%', 'Target-POC%', 'POC%', 'Realization%', 
        'Productivity', 'SOB Utilization%', 'URN', 'Status'
    ]

    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=desired_columns)
        df.to_excel(EXCEL_FILE, index=False)
        
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name='MTD Trend', dtype={'CZ ID': str})
    except Exception:
        df = pd.read_excel(EXCEL_FILE, dtype={'CZ ID': str})
    
    df.columns = [str(col).strip() for col in df.columns]
    
    rename_map = {}
    for col in df.columns:
        col_lower = col.lower()
        if col_lower in ['name', 'agent name', 'agent_name', 'names']:
            rename_map[col] = 'Names'
        elif 'cz' in col_lower and col != 'CZ ID':
            rename_map[col] = 'CZ ID'
        elif 'tl' in col_lower and col != 'TL':
            rename_map[col] = 'TL'
        elif 'status' in col_lower and col != 'Status':
            rename_map[col] = 'Status'
            
    df = df.rename(columns=rename_map)

    for col in desired_columns:
        if col not in df.columns:
            df[col] = ''
            
    # Default sabhi ko Active rakhein agar Status column khali ho
    if 'Status' in df.columns:
        df['Status'] = df['Status'].fillna('Active').replace('', 'Active')

    df = df[desired_columns]
    for col in df.select_dtypes(include=['object']):
        df[col] = df[col].fillna('')
        
    return df

def save_data(df):
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def index():
    return render_template('index.html')

# Agent Login & Performance View
@app.route('/check_performance', methods=['POST'])
def check_performance():
    cz_id = request.form.get('cz_id', '').strip()
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    
    agent = df[df['CZ ID'] == cz_id]
    if agent.empty:
        return render_template('index.html', error="Invalid CZ ID or Not Found.")
    
    agent_data = agent.iloc[0].to_dict()
    
    # Agar status Inactive hai toh access block kar do
    if str(agent_data.get('Status', 'Active')).strip().lower() == 'inactive':
        return render_template('index.html', warning="Your CZ ID is temporarily locked/inactive. Please contact the Administrator.")
        
    return render_template('profile.html', agent=agent_data)

@app.route('/admin_login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        pwd = request.form.get('password')
        if pwd == 'Anmol@#9876':
            session['admin_logged'] = True
            return redirect(url_for('admin_panel'))
        else:
            return render_template('admin_login.html', error="Incorrect Password!")
    return render_template('admin_login.html')

# Admin Panel: List all agents, upload new excel, search
@app.route('/admin_panel', methods=['GET', 'POST'])
def admin_panel():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    
    search = request.args.get('search', '').strip()
    if search:
        df = df[df['CZ ID'].astype(str).str.contains(search, case=False, na=False) | df['Names'].astype(str).str.contains(search, case=False, na=False)]
    
    agents = df.to_dict(orient='records')
    return render_template('admin_panel.html', agents=agents, search=search)

# Admin: Upload new Excel (replaces old data)
@app.route('/admin/upload_excel', methods=['POST'])
def upload_excel():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    if 'excel_file' in request.files:
        file = request.files['excel_file']
        if file.filename != '':
            if os.path.exists(EXCEL_FILE):
                os.remove(EXCEL_FILE)
            file.save(EXCEL_FILE)
            
    return redirect(url_for('admin_panel'))

# Admin: Toggle Status (Active / Inactive)
@app.route('/admin/toggle_status/<cz_id>', methods=['POST'])
def toggle_status(cz_id):
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    
    if cz_id in df['CZ ID'].values:
        idx = df[df['CZ ID'] == cz_id].index[0]
        curr = str(df.at[idx, 'Status']).strip()
        df.at[idx, 'Status'] = 'Inactive' if curr == 'Active' else 'Active'
        save_data(df)
        
    return redirect(url_for('admin_panel'))

# Admin: View & Edit Specific Agent Profile/Performance
@app.route('/admin/edit_agent/<cz_id>', methods=['GET', 'POST'])
def edit_agent(cz_id):
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    
    agent = df[df['CZ ID'] == cz_id]
    if agent.empty:
        return redirect(url_for('admin_panel'))
    
    if request.method == 'POST':
        idx = df[df['CZ ID'] == cz_id].index[0]
        for col in df.columns:
            if col in request.form:
                df.at[idx, col] = request.form.get(col)
        save_data(df)
        return redirect(url_for('admin_panel'))
        
    agent_data = agent.iloc[0].to_dict()
    return render_template('admin_edit.html', agent=agent_data)

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_logged', None)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, port=5000)
