from flask import Flask, render_template, jsonify, request, redirect, url_for, session
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'student-performance-secret-key')
BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))

SEMESTER_YEAR = {
    'Semester 1':'1st Year','Semester 2':'1st Year',
    'Semester 3':'2nd Year','Semester 4':'2nd Year',
    'Semester 5':'3rd Year','Semester 6':'3rd Year'
}
YEAR_SEMESTERS = {
    '1st Year':['Semester 1','Semester 2'],
    '2nd Year':['Semester 3','Semester 4'],
    '3rd Year':['Semester 5','Semester 6']
}
SEMESTER_FILES = {f'Semester {i}':f'semester_{i}.xlsx' for i in range(1,7)}
SUBJECTS = {
    'Semester 1':['Basic Mathematics','Communication Skills','Basic Science'],
    'Semester 2':['Applied Mathematics','Basic Electrical and Engineering','Programming in C'],
    'Semester 3':['Object Oriented Programming','Data Structure','Digital Techniques','Database Management System'],
    'Semester 4':['Java Programming','Data Communication and Network','Microprocessor Programming','Environmental Education And Sustanability'],
    'Semester 5':['Software Engineering','Opreting System','Data Analytics'],
    'Semester 6':['Mobile Application Development','Machine Learning','Software Testing','Mangement']
}
LOGIN_USERNAME = os.environ.get('ADMIN_USERNAME','silicon')
LOGIN_PASSWORD = os.environ.get('ADMIN_PASSWORD','patil')

def valid_semester(s): return s in SEMESTER_FILES
def excel_path(s): return os.path.join(BASE_FOLDER, SEMESTER_FILES[s])
def logged_in(): return session.get('logged_in') is True

def api_login_check():
    if not logged_in(): return jsonify({'success':False,'message':'Login required.'}),401
    return None

def calculate_grade(p):
    try: p=float(p)
    except: return 'F'
    if p>=90:return 'A+'
    if p>=80:return 'A'
    if p>=70:return 'B+'
    if p>=60:return 'B'
    if p>=50:return 'C'
    if p>=40:return 'D'
    return 'F'

def read_excel(s):
    if not valid_semester(s): return pd.DataFrame()
    path=excel_path(s)
    if not os.path.exists(path): return pd.DataFrame()
    try:
        df=pd.read_excel(path)
        df.columns=df.columns.astype(str).str.strip()
        return df
    except Exception as e:
        print('Excel read error:',e); return pd.DataFrame()

def process_data(df,s):
    df=df.copy()
    for c in ['Student_ID','Name','Gender','Class','Attendance']:
        if c not in df.columns: df[c]=0 if c=='Attendance' else ''
    df['Student_ID']=df['Student_ID'].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip()
    for c in ['Name','Gender','Class']:
        df[c]=df[c].fillna('').astype(str).str.strip()
    raw=df['Attendance'].fillna('').astype(str).str.strip()
    percent=raw.str.endswith('%')
    df['Attendance']=pd.to_numeric(raw.str.replace('%','',regex=False),errors='coerce')
    frac=df['Attendance'].notna() & df['Attendance'].between(0,1) & ~percent
    df.loc[frac,'Attendance']=df.loc[frac,'Attendance']*100
    df['Attendance']=df['Attendance'].fillna(0).clip(0,100).round(2)
    subs=SUBJECTS[s]
    for c in subs:
        if c not in df.columns: df[c]=0.0
        df[c]=pd.to_numeric(df[c],errors='coerce').fillna(0).clip(0,100).round(2)
    df['Total']=df[subs].sum(axis=1).round(2)
    df['Percentage']=(df['Total']/(len(subs)*100)*100).round(2)
    df['Attendance Status']=df['Attendance'].apply(lambda x:'Good' if float(x)>=75 else 'Bad')
    df['Grade']=df['Percentage'].apply(calculate_grade)
    return df

def save_excel(df,s):
    p=process_data(df,s)
    cols=['Student_ID','Name','Gender','Class']+SUBJECTS[s]+['Attendance']
    for c in cols:
        if c not in p.columns: p[c]=0 if c=='Attendance' else ''
    p[cols].to_excel(excel_path(s),index=False)

@app.route('/login',methods=['GET','POST'])
def login():
    if logged_in(): return redirect(url_for('home'))
    if request.method=='POST':
        if request.form.get('username','').strip()==LOGIN_USERNAME and request.form.get('password','')==LOGIN_PASSWORD:
            session.clear(); session['logged_in']=True; return redirect(url_for('home'))
        return render_template('index.html',login_page=True,login_error='Invalid username or password.')
    return render_template('index.html',login_page=True)

@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/')
def home():
    if not logged_in(): return redirect(url_for('login'))
    return render_template('index.html',login_page=False)

@app.route('/api/subjects')
def api_subjects():
    check=api_login_check()
    if check:return check
    s=request.args.get('semester','Semester 1').strip()
    if not valid_semester(s):return jsonify({'success':False}),400
    return jsonify({'success':True,'year':SEMESTER_YEAR[s],'semester':s,'subjects':[{'code':x,'name':x} for x in SUBJECTS[s]]})

@app.route('/api/students')
def api_students():
    check=api_login_check()
    if check:return check
    s=request.args.get('semester','Semester 1').strip(); df=read_excel(s)
    if df.empty:return jsonify([])
    df=process_data(df,s)
    cols=['Student_ID','Name','Gender','Class']+SUBJECTS[s]+['Total','Percentage','Attendance','Attendance Status','Grade']
    return jsonify(df[cols].fillna('').to_dict('records'))

def analytics_for(s):
    df=read_excel(s)
    if df.empty:return {'subjects':{},'grades':{},'students':0,'percentage':0,'attendance':0,'top':'-'}
    df=process_data(df,s)
    top='-'
    if not df.empty:
        try: top=str(df.loc[df['Percentage'].idxmax(),'Name'])
        except: pass
    return {'subjects':{x:round(float(df[x].mean()),2) for x in SUBJECTS[s]},'grades':df['Grade'].value_counts().to_dict(),'students':len(df),'percentage':round(float(df['Percentage'].mean()),2),'attendance':round(float(df['Attendance'].mean()),2),'top':top}

@app.route('/api/analytics')
def api_analytics():
    check=api_login_check()
    if check:return check
    s=request.args.get('semester','Semester 1').strip(); a=analytics_for(s)
    return jsonify({'total_students':a['students'],'average_percentage':a['percentage'],'top_performer':a['top'],'average_attendance':a['attendance'],'subjects':a['subjects'],'grades':a['grades']})

@app.route('/api/combined_analytics')
def combined_analytics():
    check=api_login_check()
    if check:return check
    year=request.args.get('year','1st Year').strip()
    if year not in YEAR_SEMESTERS:return jsonify({'success':False}),400
    return jsonify({'success':True,'year':year,'semesters':{s:analytics_for(s) for s in YEAR_SEMESTERS[year]}})

@app.route('/api/upload_excel',methods=['POST'])
def upload_excel():
    check=api_login_check()
    if check:return check
    try:
        s=request.form.get('semester','').strip(); f=request.files.get('file')
        if not valid_semester(s):return jsonify({'success':False,'message':'Invalid semester.'}),400
        if not f or not f.filename.lower().endswith(('.xlsx','.xls')):return jsonify({'success':False,'message':'Please select an Excel file.'}),400
        df=pd.read_excel(f); df.columns=df.columns.astype(str).str.strip()
        for c in ['Student_ID','Name']:
            if c not in df.columns:return jsonify({'success':False,'message':f'{c} column is required.'}),400
        save_excel(df,s)
        return jsonify({'success':True,'message':f'{s} Excel uploaded successfully.'})
    except Exception as e:return jsonify({'success':False,'message':str(e)}),500

@app.route('/api/add_student',methods=['POST'])
def add_student():
    check=api_login_check()
    if check:return check
    try:
        d=request.get_json() or {}; s=str(d.get('semester','Semester 1')).strip(); sid=str(d.get('Student_ID','')).strip(); name=str(d.get('Name','')).strip()
        if not valid_semester(s):return jsonify({'success':False,'message':'Invalid semester.'}),400
        if not sid or not name:return jsonify({'success':False,'message':'Student ID and Name are required.'}),400
        df=read_excel(s)
        if not df.empty and 'Student_ID' in df.columns:
            ids=df['Student_ID'].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip().tolist()
            if sid in ids:return jsonify({'success':False,'message':'Student ID already exists.'}),400
        row={'Student_ID':sid,'Name':name,'Gender':d.get('Gender',''),'Class':d.get('Class',''),'Attendance':d.get('Attendance',0)}
        for x in SUBJECTS[s]:row[x]=d.get(x,0)
        save_excel(pd.concat([df,pd.DataFrame([row])],ignore_index=True),s)
        return jsonify({'success':True,'message':'Student added successfully.'})
    except Exception as e:return jsonify({'success':False,'message':str(e)}),500

@app.route('/api/edit_student',methods=['POST'])
def edit_student():
    check=api_login_check()
    if check:return check
    try:
        d=request.get_json() or {}; s=str(d.get('semester','')).strip(); sid=str(d.get('Student_ID','')).strip()
        if not valid_semester(s):return jsonify({'success':False,'message':'Invalid semester.'}),400
        df=read_excel(s)
        if df.empty or 'Student_ID' not in df.columns:return jsonify({'success':False,'message':'Student not found.'}),404
        df['Student_ID']=df['Student_ID'].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip()
        matches=df.index[df['Student_ID']==sid].tolist()
        if not matches:return jsonify({'success':False,'message':'Student not found.'}),404
        i=matches[0]
        for c in ['Name','Gender','Class']:
            if c in d:df.loc[i,c]=str(d[c]).strip()
        for c in SUBJECTS[s]+['Attendance']:
            if c not in d:continue
            try:v=float(d[c]) if str(d[c]).strip() else 0.0
            except:return jsonify({'success':False,'message':f'Invalid numeric value for {c}.'}),400
            v=max(0,min(100,v))
            df[c]=pd.to_numeric(df[c],errors='coerce').fillna(0).astype(float)
            df.loc[i,c]=v
        save_excel(df,s)
        return jsonify({'success':True,'message':'Student updated successfully.'})
    except Exception as e:
        print('EDIT ERROR:',repr(e)); return jsonify({'success':False,'message':f'Update failed: {e}'}),500

@app.route('/api/delete_student',methods=['POST'])
def delete_student():
    check=api_login_check()
    if check:return check
    try:
        d=request.get_json() or {}; s=str(d.get('semester','')).strip(); sid=str(d.get('Student_ID','')).strip(); df=read_excel(s)
        if df.empty:return jsonify({'success':False,'message':'Student not found.'})
        df['Student_ID']=df['Student_ID'].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip(); new=df[df['Student_ID']!=sid].copy()
        if len(new)==len(df):return jsonify({'success':False,'message':'Student not found.'})
        save_excel(new,s); return jsonify({'success':True,'message':'Student deleted successfully.'})
    except Exception as e:return jsonify({'success':False,'message':str(e)}),500

@app.route('/health')
def health():return jsonify({'status':'ok','application':'Student Performance Analysis System'})

if __name__=='__main__':
    app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)))
