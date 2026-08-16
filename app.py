from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'student-performance-2026-secret')
BASE = os.path.dirname(os.path.abspath(__file__))

FILES = {f'Semester {i}': f'semester_{i}.xlsx' for i in range(1,7)}
YEARS = {'Semester 1':'1st Year','Semester 2':'1st Year','Semester 3':'2nd Year','Semester 4':'2nd Year','Semester 5':'3rd Year','Semester 6':'3rd Year'}
SUBJECTS = {
'Semester 1':['Basic Mathematics','Communication Skills','Engineering Physics','Engineering Chemistry','Basic Science'],
'Semester 2':['Applied Mathematics','Engineering Graphics','Basic Electrical Engineering','Programming in C','Web Page Design'],
'Semester 3':['Object Oriented Programming','Data Structure','Database Management System','Computer Networks','Operating System'],
'Semester 4':['Java Programming','Data Communication and Network','Microprocessor','Software Engineering','Python Programming'],
'Semester 5':['Advanced Java','Web Based Application Development','Software Testing','Computer Security','Project Management'],
'Semester 6':['Mobile Application Development','Cloud Computing','Artificial Intelligence','Internet of Things','Major Project']}
USER=os.environ.get('ADMIN_USERNAME','silicon')
PASS=os.environ.get('ADMIN_PASSWORD','patil')

def path(s): return os.path.join(BASE, FILES[s])
def valid(s): return s in FILES

def grade(p):
    try:p=float(p)
    except:return 'F'
    return 'A+' if p>=90 else 'A' if p>=80 else 'B+' if p>=70 else 'B' if p>=60 else 'C' if p>=50 else 'D' if p>=40 else 'F'

def read(s):
    if not valid(s) or not os.path.exists(path(s)): return pd.DataFrame()
    try:
        d=pd.read_excel(path(s)); d.columns=d.columns.astype(str).str.strip(); return d
    except Exception as e:
        print(e); return pd.DataFrame()

def process(d,s):
    d=d.copy(); subs=SUBJECTS[s]
    for c in ['Student_ID','Name','Gender','Class','Attendance']:
        if c not in d.columns:d[c]=0 if c=='Attendance' else ''
    d['Student_ID']=d['Student_ID'].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip()
    d['Name']=d['Name'].fillna('').astype(str).str.strip()
    d['Attendance']=pd.to_numeric(d['Attendance'],errors='coerce').fillna(0).clip(0,100).round(2)
    for c in subs:
        if c not in d.columns:d[c]=0
        d[c]=pd.to_numeric(d[c],errors='coerce').fillna(0).clip(0,100).round(2)
    d['Total']=d[subs].sum(axis=1).round(2)
    d['Percentage']=(d['Total']/(len(subs)*100)*100).round(2)
    d['Attendance Status']=d['Attendance'].apply(lambda x:'Good' if x>=75 else 'Bad')
    d['Grade']=d['Percentage'].apply(grade)
    return d

def save(d,s):
    d=process(d,s)
    cols=['Student_ID','Name','Gender','Class']+SUBJECTS[s]+['Attendance']
    for c in cols:
        if c not in d.columns:d[c]=''
    d[cols].to_excel(path(s),index=False)

def logged(): return session.get('logged_in') is True

@app.before_request
def guard():
    if request.endpoint in {'login','static'}: return
    if not logged():
        return jsonify({'success':False,'message':'Login required.'}),401 if request.path.startswith('/api/') else redirect(url_for('login'))

@app.route('/login',methods=['GET','POST'])
def login():
    if request.method=='POST':
        if request.form.get('username','').strip()==USER and request.form.get('password','')==PASS:
            session.clear(); session['logged_in']=True
            return redirect(url_for('home'))
        return render_template('index.html',login_page=True,login_error='Invalid username or password.')
    return render_template('index.html',login_page=True)

@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))
@app.route('/')
def home(): return render_template('index.html',login_page=False)

@app.route('/api/subjects')
def subjects():
    s=request.args.get('semester','Semester 1')
    return jsonify({'success':True,'branch':'Computer Engineering','year':YEARS[s],'semester':s,'subjects':[{'code':x,'name':x} for x in SUBJECTS[s]]}) if valid(s) else (jsonify({'success':False}),400)

@app.route('/api/students')
def students():
    s=request.args.get('semester','Semester 1'); d=read(s)
    if d.empty:return jsonify([])
    d=process(d,s); cols=['Student_ID','Name','Gender','Class']+SUBJECTS[s]+['Total','Percentage','Attendance','Attendance Status','Grade']
    return jsonify(d[cols].fillna('').to_dict('records'))

@app.route('/api/analytics')
def analytics():
    s=request.args.get('semester','Semester 1'); d=read(s)
    if d.empty:return jsonify({'total_students':0,'average_percentage':0,'top_performer':'-','average_attendance':0,'subjects':{},'grades':{}})
    d=process(d,s); top=str(d.loc[d['Percentage'].idxmax(),'Name']) if len(d) else '-'
    return jsonify({'total_students':len(d),'average_percentage':round(d['Percentage'].mean(),2),'top_performer':top,'average_attendance':round(d['Attendance'].mean(),2),'subjects':{c:round(d[c].mean(),2) for c in SUBJECTS[s]},'grades':d['Grade'].value_counts().to_dict()})

@app.route('/api/upload_excel',methods=['POST'])
def upload():
    try:
        s=request.form.get('semester','Semester 1'); f=request.files.get('file')
        if not valid(s) or not f or not f.filename.lower().endswith(('.xlsx','.xls')): return jsonify({'success':False,'message':'Valid Excel file select करा.'})
        d=pd.read_excel(f); d.columns=d.columns.astype(str).str.strip()
        if 'Student_ID' not in d.columns or 'Name' not in d.columns:return jsonify({'success':False,'message':'Student_ID आणि Name columns required आहेत.'})
        save(d,s); return jsonify({'success':True,'message':f'{s} Excel successfully saved.'})
    except Exception as e:return jsonify({'success':False,'message':str(e)})

@app.route('/api/add_student',methods=['POST'])
def add():
    try:
        x=request.get_json() or {}; s=x.get('semester','Semester 1'); d=read(s); sid=str(x.get('Student_ID','')).strip(); name=str(x.get('Name','')).strip()
        if not sid or not name:return jsonify({'success':False,'message':'Student ID आणि Name required आहेत.'})
        if not d.empty and sid in d['Student_ID'].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip().tolist():return jsonify({'success':False,'message':'Student ID already exists.'})
        row={c:x.get(c,'') for c in ['Student_ID','Name','Gender','Class','Attendance']+SUBJECTS[s]}; d=pd.concat([d,pd.DataFrame([row])],ignore_index=True); save(d,s)
        return jsonify({'success':True,'message':'Student added and Excel updated.'})
    except Exception as e:return jsonify({'success':False,'message':str(e)})

@app.route('/api/edit_student',methods=['POST'])
def edit():
    try:
        x=request.get_json() or {}; s=x.get('semester','Semester 1'); sid=str(x.get('Student_ID','')).strip(); d=read(s)
        if d.empty:return jsonify({'success':False,'message':'Student not found.'})
        d['Student_ID']=d['Student_ID'].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip(); m=d.index[d['Student_ID']==sid].tolist()
        if not m:return jsonify({'success':False,'message':'Student not found.'})
        for c in ['Name','Gender','Class','Attendance']+SUBJECTS[s]:
            if c in x:d.loc[m[0],c]=x[c]
        save(d,s); return jsonify({'success':True,'message':'Student updated in Excel.'})
    except Exception as e:return jsonify({'success':False,'message':str(e)})

@app.route('/api/delete_student',methods=['POST'])
def delete():
    try:
        x=request.get_json() or {}; s=x.get('semester','Semester 1'); sid=str(x.get('Student_ID','')).strip(); d=read(s)
        if d.empty:return jsonify({'success':False,'message':'Student not found.'})
        d['Student_ID']=d['Student_ID'].fillna('').astype(str).str.replace(r'\.0$','',regex=True).str.strip(); n=d[d['Student_ID']!=sid]
        if len(n)==len(d):return jsonify({'success':False,'message':'Student not found.'})
        save(n,s); return jsonify({'success':True,'message':'Student deleted from Excel.'})
    except Exception as e:return jsonify({'success':False,'message':str(e)})

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)),debug=False)
