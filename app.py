import os
import pandas as pd
import requests
from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = 'asr_secret_key_12345'
EXCEL_FILE = 'asr_tracker.xlsx'

# Aapka OneDrive ka direct download link
ONEDRIVE_DIRECT_URL = "https://interactaibpo1-my.sharepoint.com/:x:/g/personal/anmol_a_interactaibpo_com/IQDVzkSorGhRTKD0jp_J-T37AUAMg1-XGjG_evtfFZj5rhY?e=MZPo3q&download=1"

def load_data():
    try:
        response = requests.get(ONEDRIVE_DIRECT_URL, timeout=20)
        if response.status_code == 200:
            with open(EXCEL_FILE, 'wb') as f:
                f.write(response.content)
    except Exception as e:
        print("OneDrive sync warning:", e)

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
    df = df.rename(columns=rename_map)

    for col in desired_columns:
        if col not in df.columns:
            df[col] = ''

    df = df[desired_columns]
    
    for col in df.select_dtypes(include=['object']):
        df[col] = df[col].fillna('')
        
    return df

def save_data(df):
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/check_performance', methods=['POST'])
def check_performance():
    cz_id = request.form.get('cz_id', '').strip()
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    
    agent = df[df['CZ ID'] == cz_id]
    if agent.empty:
        return render_template('index.html', error="Invalid CZ ID or Not Found.")
    
    agent_data = agent.iloc[0].to_dict()
    
    if str(agent_data.get('Status', 'Active')).strip() == 'Inactive':
        return render_template('index.html', warning="Your CZ ID will be temporary locked.. Please contact to the Administrator.")
        
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

@app.route('/admin/update_agent', methods=['POST'])
def update_agent():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    cz_id = request.form.get('cz_id')
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    
    if cz_id in df['CZ ID'].values:
        idx = df[df['CZ ID'] == cz_id].index[0]
        for col in df.columns:
            if col != 'CZ ID' and col in request.form:
                df.at[idx, col] = request.form.get(col)
        save_data(df)
    return redirect(url_for('admin_panel'))

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_logged', None)
    return redirect(url_for('index'))

if __name__ == 'main__':
    app.run(debug=True, port=5000)
