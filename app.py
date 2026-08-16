from flask import Flask, render_template, jsonify, request, redirect, url_for, session
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'student-performance-secret-key')
BASE = os.path.dirname(os.path.abspath(__file__))

SEMESTER_FILES = {f'Semester {i}': f'semester_{i}.xlsx' for i in range(1,7)}
SEMESTER_YEAR = {
    'Semester 1':'1st Year','Semester 2':'1st Year',
    'Semester 3':'2nd Year','Semester 4':'2nd Year',
    'Semester 5':'3rd Year','Semester 6':'3rd Year'
}
SEMESTERS = {'1st Year':['Semester 1','Semester 2'],'2nd Year':['Semester 3','Semester 4'],'3rd Year':['Semester 5','Semester 6']}
SUBJECTS = {
'Semester 1':['Basic Mathematics','Communication Skills','Engineering Physics','Engineering Chemistry','Basic Science'],
'Semester 2':['Applied Mathematics','Engineering Graphics','Basic Electrical Engineering','Programming in C','Web Page Design'],
'Semester 3':['Object Oriented Programming','Data Structure','Database Management System','Computer Networks','Operating System'],
'Semester 4':['Java Programming','Data Communication and Network','Microprocessor','Software Engineering','Python Programming'],
'Semester 5':['Advanced Java','Web Based Application Development','Software Testing','Computer Security','Project Management'],
'Semester 6':['Mobile Application Development','Cloud Computing','Artificial Intelligence','Internet of Things','Major Project']}
USERNAME = os.environ.get('ADMIN_USERNAME','silicon')
PASSWORD = os.environ.get('ADMIN_PASSWORD','patil')

def valid_sem(s): return s in SEMESTER_FILES
def path(s): return os.path.join(BASE, SEMESTER_FILES[s])
def logged(): return session.get('logged_in') is True

def grade(p):
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
    if not valid_sem(s) or not os.path.exists(path(s)): return pd.DataFrame()
    try:
        df=pd.read_excel(path(s)); df.columns=df.columns.astype(str).str.strip(); return df
    except Exception as e:
        print('Excel read error:',e); return pd.DataFrame()

def process(df,s):
    df=df.copy(); subs=SUBJECTS[s]
    for c in ['Student_ID','Name','Gender','Class']:
        if c not in df.columns: df[c]=''
        df[c]=df[c].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip()
    if 'Attendance' not in df.columns: df['Attendance']=0
    df['Attendance']=pd.to_numeric(df['Attendance'],errors='coerce').fillna(0).clip(0,100).round(2)
    for c in subs:
        if c not in df.columns: df[c]=0
        df[c]=pd.to_numeric(df[c],errors='coerce').fillna(0).clip(0,100).round(2)
    df['Total']=df[subs].sum(axis=1).round(2)
    df['Percentage']=(df['Total']/(len(subs)*100)*100).round(2)
    df['Attendance Status']=df['Attendance'].apply(lambda x:'Good' if float(x)>=75 else 'Bad')
    df['Grade']=df['Percentage'].apply(grade)
    return df

def save_excel(df,s):
    df=process(df,s)
    cols=['Student_ID','Name','Gender','Class']+SUBJECTS[s]+['Attendance']
    for c in cols:
        if c not in df.columns: df[c]=''
    df[cols].to_excel(path(s),index=False)

def api_auth():
    if not logged(): return jsonify({'success':False,'message':'Login required.'}),401
    return None

@app.route('/login',methods=['GET','POST'])
def login():
    if logged(): return redirect(url_for('home'))
    error=None
    if request.method=='POST':
        u=request.form.get('username','').strip(); p=request.form.get('password','')
        if u==USERNAME and p==PASSWORD:
            session.clear(); session['logged_in']=True; return redirect(url_for('home'))
        error='Invalid username or password.'
    return render_template('index.html',login_page=True,login_error=error)

@app.route('/logout')
def logout():
    session.clear(); return redirect(url_for('login'))

@app.route('/')
def home():
    if not logged(): return redirect(url_for('login'))
    return render_template('index.html',login_page=False)

@app.route('/api/subjects')
def api_subjects():
    a=api_auth()
    if a:return a
    s=request.args.get('semester','Semester 1')
    if not valid_sem(s): return jsonify({'success':False,'message':'Invalid semester.'}),400
    return jsonify({'success':True,'branch':'Computer Engineering','year':SEMESTER_YEAR[s],'semester':s,'subjects':[{'code':x,'name':x} for x in SUBJECTS[s]]})

@app.route('/api/students')
def api_students():
    a=api_auth()
    if a:return a
    s=request.args.get('semester','Semester 1')
    if not valid_sem(s):return jsonify([])
    df=read_excel(s)
    if df.empty:return jsonify([])
    df=process(df,s)
    cols=['Student_ID','Name','Gender','Class']+SUBJECTS[s]+['Total','Percentage','Attendance','Attendance Status','Grade']
    return jsonify(df[cols].fillna('').to_dict('records'))

@app.route('/api/analytics')
def analytics():
    a=api_auth()
    if a:return a
    s=request.args.get('semester','Semester 1'); df=read_excel(s)
    if df.empty:return jsonify({'total_students':0,'average_percentage':0,'top_performer':'-','average_attendance':0,'subjects':{},'grades':{}})
    df=process(df,s); top=df.loc[df['Percentage'].idxmax(),'Name'] if not df.empty else '-'
    return jsonify({'total_students':int(len(df)),'average_percentage':round(float(df['Percentage'].mean()),2),'top_performer':str(top),'average_attendance':round(float(df['Attendance'].mean()),2),'subjects':{x:round(float(df[x].mean()),2) for x in SUBJECTS[s]},'grades':df['Grade'].value_counts().to_dict()})

@app.route('/api/upload_excel',methods=['POST'])
def upload_excel():
    a=api_auth()
    if a:return a
    try:
        s=request.form.get('semester','').strip(); f=request.files.get('file')
        if not valid_sem(s):return jsonify({'success':False,'message':'Invalid semester selected.'}),400
        if not f or not f.filename:return jsonify({'success':False,'message':'Please select Excel file.'}),400
        if not f.filename.lower().endswith(('.xlsx','.xls')):return jsonify({'success':False,'message':'Only Excel files are allowed.'}),400
        df=pd.read_excel(f)
        if df.empty:return jsonify({'success':False,'message':'Uploaded Excel is empty.'}),400
        df.columns=df.columns.astype(str).str.strip()
        if 'Student_ID' not in df.columns:return jsonify({'success':False,'message':'Student_ID column is required.'}),400
        if 'Name' not in df.columns:return jsonify({'success':False,'message':'Name column is required.'}),400
        save_excel(df,s); return jsonify({'success':True,'message':f'{s} Excel uploaded successfully.'})
    except Exception as e:return jsonify({'success':False,'message':f'Excel upload failed: {e}'}),500

def get_data():
    d=request.get_json() or {}; s=str(d.get('semester','Semester 1')).strip(); return d,s

@app.route('/api/add_student',methods=['POST'])
def add_student():
    a=api_auth()
    if a:return a
    try:
        d,s=get_data(); sid=str(d.get('Student_ID','')).strip(); name=str(d.get('Name','')).strip()
        if not valid_sem(s):return jsonify({'success':False,'message':'Invalid semester.'}),400
        if not sid or not name:return jsonify({'success':False,'message':'Student ID and Name are required.'}),400
        df=read_excel(s)
        if not df.empty:
            ids=df.get('Student_ID',pd.Series(dtype=str)).fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip()
            if sid in ids.tolist():return jsonify({'success':False,'message':'Student ID already exists.'})
        row={'Student_ID':sid,'Name':name,'Gender':d.get('Gender',''),'Class':d.get('Class',''),'Attendance':d.get('Attendance',0)}
        row.update({x:d.get(x,0) for x in SUBJECTS[s]}); df=pd.concat([df,pd.DataFrame([row])],ignore_index=True); save_excel(df,s)
        return jsonify({'success':True,'message':'Student added successfully.'})
    except Exception as e:return jsonify({'success':False,'message':str(e)}),500

@app.route('/api/edit_student',methods=['POST'])
def edit_student():
    a=api_auth()
    if a:return a
    try:
        d,s=get_data(); sid=str(d.get('Student_ID','')).strip(); df=read_excel(s)
        if df.empty or 'Student_ID' not in df.columns:return jsonify({'success':False,'message':'Student not found.'})
        ids=df['Student_ID'].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip(); m=df.index[ids==sid].tolist()
        if not m:return jsonify({'success':False,'message':'Student not found.'})
        i=m[0]
        for c in ['Name','Gender','Class','Attendance']+SUBJECTS[s]:
            if c in d:df.loc[i,c]=d[c]
        save_excel(df,s);return jsonify({'success':True,'message':'Student updated successfully.'})
    except Exception as e:return jsonify({'success':False,'message':str(e)}),500

@app.route('/api/delete_student',methods=['POST'])
def delete_student():
    a=api_auth()
    if a:return a
    try:
        d,s=get_data();sid=str(d.get('Student_ID','')).strip();df=read_excel(s)
        if df.empty:return jsonify({'success':False,'message':'Student not found.'})
        ids=df['Student_ID'].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip(); new=df[ids!=sid].copy()
        if len(new)==len(df):return jsonify({'success':False,'message':'Student not found.'})
        save_excel(new,s);return jsonify({'success':True,'message':'Student deleted successfully.'})
    except Exception as e:return jsonify({'success':False,'message':str(e)}),500

@app.route('/health')
def health():return jsonify({'status':'ok','application':'Student Performance Analysis System'})

if __name__=='__main__':app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)))
