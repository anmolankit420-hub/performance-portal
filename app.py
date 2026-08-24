import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file

app = Flask(__name__)
app.secret_key = 'asr_secret_key_12345'
EXCEL_FILE = 'asr_tracker.xlsx'

def get_initial_data():
    initial_agents = [
        ("41368", "Anjali Kumari"), ("41414", "Shilpa Kumari"), ("44624", "Anita Kujur"),
        ("45466", "Laxmi Kumari"), ("45552", "Nisha Kumari"), ("45755", "Dimpy Kumari"),
        ("45808", "Pooja Kumari"), ("46802", "Shivani Kumari"), ("47877", "Shivani Kumari"),
        ("47314", "Pinki Kumari"), ("451993", "Priti Kumari"), ("47227", "Komal Kumari"),
        ("47235", "Rakhi Kumari"), ("47586", "Md. Tosif Raja"), ("47595", "Kajal kumari"),
        ("300406", "khushboo Kumari Sharma"), ("48193", "Nilwanti Kumari"), ("48706", "Vinita Kumari"),
        ("48724", "Babita Kumari"), ("48867", "Deepak Hans"), ("49201", "Saein Kumar Singh"),
        ("49332", "Rose Toppo"), ("404610", "Aaliya Perween"), ("450262", "Gulnaz Perween"),
        ("450272", "Nitu Kumari"), ("450312", "Sujata Kumari"), ("450317", "Krity Kumari"),
        ("450338", "Rim Reyan Fatima"), ("450420", "Chandni Goswami"), ("450456", "Anju Hembrom"),
        ("450550", "Shaheen Parween"), ("450708", "Nidhi Kumari"), ("450714", "Princi Kumari"),
        ("450718", "Poonam Hembrom"), ("450722", "Juhi Kumari"), ("450740", "Pramila Manjhi"),
        ("450746", "Sazia Parween"), ("450751", "Saba Parween"), ("450938", "Rafat Parween"),
        ("450941", "Shweta Kumari"), ("450945", "Renu Toppo"), ("450974", "Rani Kumari"),
        ("450985", "Ritesh Kumar"), ("451502", "Chandan Kumar Mahto"), ("451818", "Yogendra Mahto"),
        ("451825", "Roshan Kumar Saw"), ("451846", "Ankit Kumar"), ("451899", "Sonu Ray"),
        ("452381", "Prabhakar Kumar"), ("452586", "Dinesh Kumar Mahto"), ("452597", "Subroto Banerjee"),
        ("300234", "Megha sharma"), ("300306", "Anshu Kumari"), ("453034", "Shanjana Gowswami"),
        ("453050", "Fiza Saba"), ("453481", "Neha Rani"), ("453488", "Beauty Gorai"),
        ("456923", "Sonali Kumari Giri"), ("456947", "Arjun Kumar"), ("40468", "Simran Saba"),
        ("31231", "Rahul Kumar Raj"), ("49959", "Neha Kumari"), ("3734", "Anjali Oraon"),
        ("200020", "Monika Singh"), ("200087", "Reemu"), ("46858", "Shagufta Parween"),
        ("454993", "Sanju Sharma"), ("454995", "Anshu Kumari"), ("455137", "Anita Kujur"),
        ("455768", "Shanti Sahu"), ("20891", "Sajia Tabassum"), ("20900", "Rinki Kumari"),
        ("4835", "Khushi Kumari"), ("4846", "Sayna Parween"), ("454916", "Shubham"),
        ("450199", "Mohd Shahid"), ("450203", "Noorish"), ("450211", "Arshad"),
        ("493642", "Sapna"), ("493643", "Satyam Kumar"), ("2096", "Khushboo kumari"),
        ("2180", "Rohit Ray"), ("2124", "Kajal kumari"), ("2126", "Rupa Kumari"),
        ("2513", "Mukta Dung dung"), ("2841", "Hemant"), ("2881", "Vishal Yadav"),
        ("494194", "Anita kumari"), ("494196", "Navneet Verma"), ("4028", "Afzal"),
        ("4123", "Sachin"), ("3518", "Rahul Kumar Ram"), ("3544", "Sanjay Kumar Mahato"),
        ("30476", "Kainat Anwar"), ("30477", "Karan Yadav"), ("30478", "Amit kumar mishra"),
        ("494102", "Yamini Hembrom"), ("494103", "Rachita kumari"), ("494105", "Tarun Das"),
        ("30565", "Rishi Sharma"), ("30596", "Raushan Kumar"), ("30602", "Madhuri kumari"),
        ("30606", "Jaya kumari dubey"), ("30607", "Shifa kainat"), ("30609", "Warka jahan"),
        ("30610", "Sahista Parween"), ("30812", "Harsh Sharma"), ("30814", "Vaibhav Kumar"),
        ("30815", "Falak Parween"), ("493439", "Sabnam Khatoon"), ("493443", "Shubham Kumar Sharma"),
        ("493445", "Md Adil Ansari"), ("493446", "Rajkumar Sinha"), ("493450", "Komal Kumari Nayak"),
        ("493452", "Jyoti Kumari"), ("492824", "Kumari Roopa"), ("492826", "Gulafsha perween"),
        ("492827", "Khushboo kumari"), ("492828", "Gulapi Soren"), ("492830", "Sanjana priya"),
        ("492833", "Nitin kaoriyar"), ("492834", "Dibyani kumari ram"), ("31259", "Simran Sajid"),
        ("31262", "Shaifali Shahid"), ("30903", "Puja Kumari Saw"), ("30904", "Sadiya perween"),
        ("30905", "Nikita kumari"), ("30922", "Sarita Kumari"), ("30923", "Bidisha Chowdhury"),
        ("30925", "Sonam perween"), ("30928", "Anjali Kumari"), ("30931", "Preety cerketta"),
        ("30942", "Nafesa perween"), ("30999", "Harish yadav"), ("31001", "Shikha Kumari"),
        ("31005", "Ananya Gupta"), ("41320", "Mariya Firdous"), ("41321", "Riya kumari"),
        ("41322", "Shaksham sahu"), ("41330", "Varsha Tirkey"), ("41332", "Deepti horo"),
        ("41340", "Riya kumari"), ("41347", "Namita kumari"), ("41350", "Jamil Akhtar"),
        ("41353", "Khushboo kumari"), ("41359", "Tania Sultana"), ("41219", "Supriya hansda"),
        ("41221", "Kajal kumari"), ("41299", "Nisha kumari ray"), ("41436", "Priyanshu Raj"),
        ("31086", "Shubham Kumar"), ("31087", "Shruti Pandey"), ("31088", "Anjali Lugun"),
        ("31089", "Khushi Singh munda"), ("31090", "Sandhya Kumari"), ("31092", "Tripti Mondal"),
        ("31094", "Farhan Khan"), ("31100", "Nasrin Khatoon"), ("31102", "Monika Runda"),
        ("31103", "Raman"), ("31104", "Anima Munda"), ("40223", "Milirani tiriya"),
        ("40231", "Sneha kumari"), ("40282", "Kumkum Kumari"), ("40303", "Priyanka kumari"),
        ("40374", "Kajal yadav"), ("40395", "Abdulla Siddique"), ("40444", "Rohit Singh"),
        ("4136", "Ankit bharti"), ("47784", "Santosh Toppo"), ("47785", "Banti Kumari"),
        ("47790", "Jyoti kumari"), ("47815", "Anju kumari"), ("47863", "Gaurav kumar"),
        ("47867", "Uday Marandi"), ("47899", "Sunita kumari"), ("47905", "Aman raza"),
        ("47911", "Md SARIK ANSARI"), ("47977", "Saniya Kumari"), ("48077", "Nupur Bid"),
        ("48098", "Sahajad Ansari"), ("48120", "Priya kumari"), ("48154", "Reema kumari"),
        ("48184", "Jaid ansari"), ("48207", "Pratima lakra")
    ]
    return initial_agents

def load_data():
    desired_columns = [
        'CZ ID', 'Names', 'Shift', 'TL', 'QA', 'PIP', 'June', 'July', 'MTD Aug', 
        'D-2', 'D-1', 'D-Day', 'D-Day SOB POC%', 'Mandays', 'CPA', 
        'Target-Booking%', 'Booking%', 'Target-POC%', 'POC%', 'Realization%', 
        'Productivity', 'SOB Utilization%', 'URN', 'Status'
    ]

    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=desired_columns)
        initial = get_initial_data()
        init_rows = []
        for cz, name in initial:
            row = {col: '' for col in desired_columns}
            row['CZ ID'] = cz
            row['Names'] = name
            row['Status'] = 'Active'
            init_rows.append(row)
        df = pd.DataFrame(init_rows)
        df.to_excel(EXCEL_FILE, index=False)
        
    try:
        # dtype=str rakhne se excel ke percentages/decimals apni exact string form mein read hote hain bina round-off ya float conversion ke
        df = pd.read_excel(EXCEL_FILE, sheet_name='MTD Trend', dtype=str)
    except Exception:
        df = pd.read_excel(EXCEL_FILE, dtype=str)
    
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
            
    if 'Status' in df.columns:
        df['Status'] = df['Status'].fillna('Active').replace('', 'Active')

    df = df[desired_columns]
    for col in df.select_dtypes(include=['object']):
        df[col] = df[col].fillna('')
        df[col] = df[col].replace('nan', '')
        df[col] = df[col].replace('None', '')
        
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
    
    if str(agent_data.get('Status', 'Active')).strip().lower() == 'inactive':
        return render_template('index.html', warning="Your CZ ID is temporarily locked/inactive. Please contact the Administrator.")
        
    return render_template('profile.html', agent=agent_data)

@app.route('/admin_login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        pwd = request.form.get('password')
        if pwd == 'Anmol@9876#':
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
    
    total_active = len(df[df['Status'].str.strip().str.lower() == 'active'])
    total_inactive = len(df[df['Status'].str.strip().str.lower() == 'inactive'])
    
    agents = df.to_dict(orient='records')
    return render_template('admin_panel.html', agents=agents, search=search, total_active=total_active, total_inactive=total_inactive)

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
            flash('Excel file uploaded and data replaced successfully!', 'success')
            
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
    
    if cz_id in df['CZ ID'].values:
        flash('CZ ID already exists!', 'error')
        return redirect(url_for('admin_panel'))
        
    new_row = {col: '' for col in df.columns}
    new_row['CZ ID'] = cz_id
    new_row['Names'] = name
    new_row['TL'] = tl
    new_row['Shift'] = shift
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
    
    df = df[df['CZ ID'] != cz_id]
    save_data(df)
    flash('Agent deleted successfully!', 'success')
    return redirect(url_for('admin_panel'))

@app.route('/admin/toggle_status/<cz_id>', methods=['POST'])
def toggle_status(cz_id):
    if not session.get('admin_logged'):
        return redirect(url_for('admin_login'))
    
    df = load_data()
    df = df.loc[:, ~df.columns.duplicated()]
    
    if cz_id in df['CZ ID'].values:
        idx = df[df['CZ ID'] == cz_id].index[0]
        curr = str(df.at[idx, 'Status']).strip()
        new_status = 'Inactive' if curr == 'Active' else 'Active'
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
    
    agent = df[df['CZ ID'] == cz_id]
    if agent.empty:
        return redirect(url_for('admin_panel'))
    
    if request.method == 'POST':
        idx = df[df['CZ ID'] == cz_id].index[0]
        for col in df.columns:
            if col in request.form:
                val = request.form.get(col)
                if val == 'nan' or val is None:
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
