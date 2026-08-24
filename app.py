import os
import pandas as pd
import requests
from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)
app.secret_key = 'asr_secret_key_12345'
EXCEL_FILE = 'asr_tracker.xlsx'

# Yahan apna OneDrive ka Direct Download Link daalein
ONEDRIVE_DIRECT_URL = "https://interactaibpo1-my.sharepoint.com/:x:/g/personal/anmol_a_interactaibpo_com/IQDVzkSorGhRTKD0jp_J-T37AUAMg1-XGjG_evtfFZj5rhY?e=MZPo3q&download=1"

def load_data():
    try:
        response = requests.get(ONEDRIVE_DIRECT_URL)
        if response.status_code == 200:
            with open(EXCEL_FILE, 'wb') as f:
                f.write(response.content)
    except Exception as e:
        print("Auto-sync warning:", e)

    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            'CZ ID', 'Name', 'TL', 'D-Day', 'D-1', 'MTD August', 'D-Day SOB POC%',  
            'Mandays', 'CPA', 'Target Booking', 'Booking', 'Booking %',  
            'Target POC%', 'POC', 'Realisation', 'Productivity', 'SOB Utilisation', 'URN', 'Status'
        ])
        df.to_excel(EXCEL_FILE, index=False)
        
    df = pd.read_excel(EXCEL_FILE, dtype={'CZ ID': str})
    
    # Sirf text/object columns ko blank se fill karein taaki float columns crash na hon
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
    
    agent = df[df['CZ ID'] == cz_id]
    if agent.empty:
        return render_template('index.html', error="Invalid CZ ID or Not Found.")
    
    agent_data = agent.iloc[0].to_dict()
    
    if agent_data.get('Status', 'Active') == 'Inactive':
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
    search = request.args.get('search', '').strip()
    if search:
        df = df[df['CZ ID'].str.contains(search, case=False) | df['Name'].str.contains(search, case=False)]
    
    agents = df.to_dict(orient='records')
    return render_template('admin_panel.html', agents=agents, search=search)

@app.route('/admin/update_agent', methods=['POST'])
def update_agent():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    cz_id = request.form.get('cz_id')
    df = load_data()
    
    if cz_id in df['CZ ID'].values:
        idx = df[df['CZ ID'] == cz_id].index[0]
        for col in df.columns:
            if col != 'CZ ID' and col in request.form:
                df.at[idx, col] = request.form.get(col)
        save_data(df)
    return redirect(url_for('admin_panel'))

@app.route('/admin/add_agent', methods=['POST'])
def add_agent():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    cz_id = request.form.get('cz_id', '').strip()
    name = request.form.get('name', '').strip()
    tl = request.form.get('tl', 'Default TL').strip()
    
    df = load_data()
    if cz_id in df['CZ ID'].values:
        return redirect(url_for('admin_panel'))
    
    new_row = {
        'CZ ID': cz_id, 'Name': name, 'TL': tl,
        'D-Day': 85.0, 'D-1': 82.0, 'MTD August': 84.5, 'D-Day SOB POC%': 20.0,
        'Mandays': 22, 'CPA': 200.0, 'Target Booking': 60, 'Booking': 45, 'Booking %': '75.0%',
        'Target POC%': '20.0%', 'POC': '25.0%', 'Realisation': 90.0, 'Productivity': 95.0, 
        'SOB Utilisation': 88.0, 'URN': 250, 'Status': 'Active'
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    save_data(df)
    return redirect(url_for('admin_panel'))

@app.route('/admin/toggle_status/<cz_id>', methods=['POST'])
def toggle_status(cz_id):
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    df = load_data()
    if cz_id in df['CZ ID'].values:
        idx = df[df['CZ ID'] == cz_id].index[0]
        curr = df.at[idx, 'Status']
        df.at[idx, 'Status'] = 'Inactive' if curr == 'Active' else 'Active'
        save_data(df)
    return redirect(url_for('admin_panel'))

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_logged', None)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, port=5000)
