import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify

app = Flask(__name__)
app.secret_key = 'asr_secret_key_12345'
EXCEL_FILE = 'asr_tracker.xlsx'
ADMIN_PASSWORD = 'Anmol@9876#'

def load_data():
    if not os.path.exists(EXCEL_FILE):
        print("Excel file not found. Creating a new sample file...")
        df = pd.DataFrame(columns=[
            'CZ ID', 'Names', 'Shift', 'TL', 'QA', 'PIP', 'June', 'July', 'MTD Aug', 
            'D-2', 'D-1', 'D-Day', 'D-Day SOB POC%', 'Mandays', 'CPA', 
            'Target-Booking%', 'Booking%', 'Target-POC%', 'POC%', 'Realization%', 
            'Productivity', 'SOB Utilization%', 'URN', 'Status'
        ])
        df.to_excel(EXCEL_FILE, index=False)
        
    try:
        temp_df = pd.read_excel(EXCEL_FILE, sheet_name=0, header=None, dtype=str)
        print(f"Excel successfully opened. Total rows found: {len(temp_df)}")
        header_row = 0
        
        for idx, row in temp_df.iterrows():
            row_str = ' '.join(row.astype(str).values).lower()
            if 'cz' in row_str and 'id' in row_str:
                header_row = idx
                print(f"Header row detected at index: {header_row}")
                break
                
        df = pd.read_excel(EXCEL_FILE, sheet_name=0, header=header_row, dtype=str)
    except Exception as e:
        print(f"CRITICAL ERROR reading excel: {e}")
        df = pd.DataFrame(columns=['CZ ID', 'Names', 'Status'])
    
    df.columns = [str(col).strip() for col in df.columns if str(col).strip() != 'nan']
    
    rename_map = {}
    for col in df.columns:
        col_lower = col.lower().replace('_', ' ').strip()
        
        if col_lower in ['name', 'agent name', 'agent_name', 'names', 'employee name']:
            rename_map[col] = 'Names'
        elif 'cz' in col_lower and 'id' in col_lower:
            rename_map[col] = 'CZ ID'
        elif col_lower in ['tl', 'team leader', 'supervisor']:
            rename_map[col] = 'TL'
        elif col_lower in ['qa', 'quality analyst']:
            rename_map[col] = 'QA'
        elif col_lower in ['shift', 'timing']:
            rename_map[col] = 'Shift'
        elif col_lower in ['status', 'state']:
            rename_map[col] = 'Status'
            
    df = df.rename(columns=rename_map)

    for col in ['CZ ID', 'Names', 'Status']:
        if col not in df.columns:
            df[col] = ''
            
    if 'Status' in df.columns:
        df['Status'] = df['Status'].fillna('Active').replace('', 'Active')

    for col in df.columns:
        df[col] = df[col].fillna('').astype(str).replace(['nan', 'None', 'NAN', '<NA>'], '')
        
    df = df[df['CZ ID'].str.strip() != '']
    print(f"Data loaded successfully. Total active agent rows: {len(df)}")
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
    
    agent = df[df['CZ ID'].astype(str).str.strip() == cz_id]
    if agent.empty:
        return render_template('index.html', error="Invalid CZ ID or Not Found.")
    
    agent_data = agent.iloc[0].to_dict()
    
    if str(agent_data.get('Status', 'Active')).strip().lower() == 'inactive':
        return render_template('index.html', warning="Your CZ ID is temporarily locked/inactive. Please contact the Administrator.")
        
    return render_template('profile.html', agent=agent_data)

@app.route('/admin_login', methods=['GET', 'POST'])
def admin_login():
    error = None
    if request.method == 'POST':
        pwd = request.form.get('password', '').strip()
        if pwd == ADMIN_PASSWORD:
            session.clear()
            session['admin_logged'] = True
            return redirect(url_for('admin_panel'))
        else:
            error = "Incorrect Password!"
    return render_template('admin_login.html', error=error)

@app.route('/admin_panel', methods=['GET', 'POST'])
def admin_panel():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    
    search = request.args.get('search', '').strip()
    if search:
        df = df[df['CZ ID'].astype(str).str.contains(search, case=False, na=False) | df['Names'].astype(str).str.contains(search, case=False, na=False)]
    
    total_active = len(df[df['Status'].str.strip().str.lower() == 'active'])
    total_inactive = len(df[df['Status'].str.strip().str.lower() == 'inactive'])
    
    # फाइल की जानकारी चेक करने के लिए
    file_exists = os.path.exists(EXCEL_FILE)
    file_name = EXCEL_FILE if file_exists else None
    
    columns = list(df.columns)
    agents = df.to_dict(orient='records')
    return render_template('admin_panel.html', agents=agents, columns=columns, search=search, total_active=total_active, total_inactive=total_inactive, file_exists=file_exists, file_name=file_name)

@app.route('/admin/upload_excel', methods=['POST'])
def upload_excel():
    if not session.get('admin_logged'):
        return jsonify({'success': False, 'error': 'Unauthorized'}), 401
    
    if 'excel_file' in request.files:
        file = request.files['excel_file']
        if file.filename != '':
            try:
                file.save(EXCEL_FILE)
                return jsonify({'success': True, 'message': 'Excel uploaded successfully!'})
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)}), 500
                
    return jsonify({'success': False, 'error': 'No file selected'}), 400

@app.route('/admin/delete_excel', methods=['POST'])
def delete_excel():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    if os.path.exists(EXCEL_FILE):
        try:
            os.remove(EXCEL_FILE)
            flash('Current Excel file deleted successfully!', 'success')
        except Exception as e:
            flash(f'Error deleting file: {e}', 'error')
    else:
        flash('No file found to delete!', 'error')
        
    return redirect(url_for('admin_panel'))

@app.route('/admin/download_excel')
def download_excel():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True)
    return redirect(url_for('admin_panel'))

@app.route('/admin/add_agent', methods=['POST'])
def add_agent():
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    cz_id = request.form.get('cz_id', '').strip()
    name = request.form.get('name', '').strip()
    tl = request.form.get('tl', '').strip()
    shift = request.form.get('shift', '').strip()
    
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    
    if cz_id in df['CZ ID'].astype(str).values:
        flash('CZ ID already exists!', 'error')
        return redirect(url_for('admin_panel'))
        
    new_row = {col: '' for col in df.columns}
    new_row['CZ ID'] = cz_id
    new_row['Names'] = name
    if 'TL' in df.columns: new_row['TL'] = tl
    if 'Shift' in df.columns: new_row['Shift'] = shift
    new_row['Status'] = 'Active'
    
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    save_data(df)
    flash(f'Agent {name} added successfully!', 'success')
    return redirect(url_for('admin_panel'))

@app.route('/admin/delete_agent/<cz_id>', methods=['POST'])
def delete_agent(cz_id):
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
        
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    df = df[df['CZ ID'].astype(str) != str(cz_id)]
    save_data(df)
    flash('Agent deleted successfully!', 'success')
    return redirect(url_for('admin_panel'))

@app.route('/admin/toggle_status/<cz_id>', methods=['POST'])
def toggle_status(cz_id):
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    matches = df[df['CZ ID'].astype(str) == str(cz_id)]
    if not matches.empty:
        idx = matches.index[0]
        curr = str(df.at[idx, 'Status']).strip()
        new_status = 'Inactive' if curr.lower() == 'active' else 'Active'
        df.at[idx, 'Status'] = new_status
        save_data(df)
        flash(f'Agent status changed to {new_status}!', 'success')
        
    return redirect(url_for('admin_panel'))

@app.route('/admin/edit_agent/<cz_id>', methods=['GET', 'POST'])
def edit_agent(cz_id):
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    agent = df[df['CZ ID'].astype(str) == str(cz_id)]
    if agent.empty:
        return redirect(url_for('admin_panel'))
    
    if request.method == 'POST':
        idx = agent.index[0]
        for col in df.columns:
            if col in request.form:
                val = request.form.get(col)
                if val in ['nan', 'None', None]:
                    val = ''
                df.loc[idx, col] = str(val)
        save_data(df)
        flash('Agent performance updated successfully!', 'success')
        return redirect(url_for('admin_panel'))
        
    agent_data = agent.iloc[0].to_dict()
    cleaned_agent_data = {k: ('' if str(v).lower() in ['nan', 'none'] else v) for k, v in agent_data.items()}
    
    return render_template('admin_edit.html', agent=cleaned_agent_data)

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_logged', None)
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, port=5000)
