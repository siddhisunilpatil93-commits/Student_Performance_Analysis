from flask import Flask, render_template, jsonify, request, redirect, url_for, session
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "student-performance-change-this-secret")
BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))

SEMESTERS = {
    "1st Year": ["Semester 1", "Semester 2"],
    "2nd Year": ["Semester 3", "Semester 4"],
    "3rd Year": ["Semester 5", "Semester 6"]
}

SEMESTER_YEAR = {
    "Semester 1": "1st Year", "Semester 2": "1st Year",
    "Semester 3": "2nd Year", "Semester 4": "2nd Year",
    "Semester 5": "3rd Year", "Semester 6": "3rd Year"
}

SEMESTER_FILES = {f"Semester {i}": f"semester_{i}.xlsx" for i in range(1, 7)}

SUBJECTS = {
    "Semester 1": ["Basic Mathematics","Communication Skills","Engineering Physics","Engineering Chemistry","Basic Science"],
    "Semester 2": ["Applied Mathematics","Engineering Graphics","Basic Electrical Engineering","Programming in C","Web Page Design"],
    "Semester 3": ["Object Oriented Programming","Data Structure","Database Management System","Computer Networks","Operating System"],
    "Semester 4": ["Java Programming","Data Communication and Network","Microprocessor","Software Engineering","Python Programming"],
    "Semester 5": ["Advanced Java","Web Based Application Development","Software Testing","Computer Security","Project Management"],
    "Semester 6": ["Mobile Application Development","Cloud Computing","Artificial Intelligence","Internet of Things","Major Project"]
}

def path_for(semester):
    return os.path.join(BASE_FOLDER, SEMESTER_FILES[semester])

def valid_semester(s):
    return s in SEMESTER_FILES

def grade(p):
    try: p=float(p)
    except: return "F"
    if p>=90:return "A+"
    if p>=80:return "A"
    if p>=70:return "B+"
    if p>=60:return "B"
    if p>=50:return "C"
    if p>=40:return "D"
    return "F"

def read_df(semester):
    if not valid_semester(semester) or not os.path.exists(path_for(semester)):
        return pd.DataFrame()
    try:
        df=pd.read_excel(path_for(semester))
        df.columns=df.columns.astype(str).str.strip()
        return df
    except Exception as e:
        print(e); return pd.DataFrame()

def process(df, semester):
    df=df.copy()
    for c in ["Student_ID","Name","Gender","Class","Attendance"]:
        if c not in df.columns: df[c]="" if c!="Attendance" else 0
    df["Student_ID"]=df["Student_ID"].astype(str).str.replace(r"\.0$","",regex=True).str.strip()
    df["Name"]=df["Name"].fillna("").astype(str).str.strip()
    df["Attendance"]=pd.to_numeric(df["Attendance"],errors="coerce").fillna(0).clip(0,100).round(2)
    subs=SUBJECTS[semester]
    for s in subs:
        if s not in df.columns: df[s]=0
        df[s]=pd.to_numeric(df[s],errors="coerce").fillna(0).clip(0,100)
    df["Total"]=df[subs].sum(axis=1).round(2)
    df["Percentage"]=(df["Total"]/(len(subs)*100)*100).round(2)
    df["Attendance Status"]=df["Attendance"].apply(lambda x:"Good" if x>=75 else "Bad")
    df["Grade"]=df["Percentage"].apply(grade)
    return df

def save_df(df, semester):
    raw=[c for c in ["Student_ID","Name","Gender","Class"]+SUBJECTS[semester]+["Attendance"] if c in df.columns]
    process(df,semester)[raw].to_excel(path_for(semester),index=False)


# =========================================================
# LOGIN / SECURITY
# =========================================================

LOGIN_USERNAME = os.environ.get("ADMIN_USERNAME", "silicon")
LOGIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "patil")


def login_required():
    return session.get("logged_in") is True


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        if username == LOGIN_USERNAME and password == LOGIN_PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("home"))

        return render_template(
            "index.html",
            login_page=True,
            login_error="Invalid username or password."
        )

    return render_template("index.html", login_page=True)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.before_request
def protect_dashboard_and_api():
    allowed = {"login", "logout", "static"}
    if request.endpoint not in allowed and not login_required():
        if request.path.startswith("/api/"):
            return jsonify({"success": False, "message": "Login required."}), 401
        return redirect(url_for("login"))


@app.route("/")
def home():
    if not login_required():
        return redirect(url_for("login"))
    return render_template("index.html", login_page=False)

@app.route("/api/subjects")
def api_subjects():
    s=request.args.get("semester","Semester 1")
    if not valid_semester(s): return jsonify({"success":False}),400
    return jsonify({"success":True,"branch":"Computer Engineering","year":SEMESTER_YEAR[s],
                    "semester":s,"subjects":[{"code":x,"name":x} for x in SUBJECTS[s]]})

@app.route("/api/students")
def api_students():
    s=request.args.get("semester","Semester 1")
    df=read_df(s)
    if df.empty:return jsonify([])
    df=process(df,s)
    cols=["Student_ID","Name","Gender","Class"]+SUBJECTS[s]+["Total","Percentage","Attendance","Attendance Status","Grade"]
    return jsonify(df[cols].fillna("").to_dict("records"))

@app.route("/api/analytics")
def api_analytics():
    s=request.args.get("semester","Semester 1"); df=read_df(s)
    if df.empty:return jsonify({"total_students":0,"average_percentage":0,"top_performer":"-","average_attendance":0,"subjects":{},"grades":{}})
    df=process(df,s)
    top=df.loc[df["Percentage"].idxmax(),"Name"] if len(df) else "-"
    return jsonify({"total_students":len(df),
                    "average_percentage":round(df["Percentage"].mean(),2),
                    "top_performer":str(top),
                    "average_attendance":round(df["Attendance"].mean(),2),
                    "subjects":{x:round(df[x].mean(),2) for x in SUBJECTS[s]},
                    "grades":df["Grade"].value_counts().to_dict()})

@app.route("/api/upload_excel",methods=["POST"])
def upload():
    try:
        s=request.form.get("semester","Semester 1")
        f=request.files.get("file")
        if not valid_semester(s) or not f or not f.filename.lower().endswith((".xlsx",".xls")):
            return jsonify({"success":False,"message":"Valid Excel file select करा."})
        df=pd.read_excel(f); df.columns=df.columns.astype(str).str.strip()
        if "Student_ID" not in df.columns or "Name" not in df.columns:
            return jsonify({"success":False,"message":"Student_ID आणि Name columns required आहेत."})
        save_df(df,s)
        return jsonify({"success":True,"message":f"{s} Excel successfully saved."})
    except Exception as e:return jsonify({"success":False,"message":str(e)})

@app.route("/api/add_student",methods=["POST"])
def add_student():
    try:
        d=request.get_json(); s=d.get("semester","Semester 1")
        df=read_df(s)
        row={c:d.get(c,"") for c in ["Student_ID","Name","Gender","Class","Attendance"]+SUBJECTS[s]}
        if not str(row["Student_ID"]).strip() or not str(row["Name"]).strip():
            return jsonify({"success":False,"message":"Student ID आणि Name required आहेत."})
        if not df.empty and str(row["Student_ID"]) in df["Student_ID"].astype(str).values:
            return jsonify({"success":False,"message":"Student ID already exists."})
        df=pd.concat([df,pd.DataFrame([row])],ignore_index=True)
        save_df(df,s)
        return jsonify({"success":True,"message":"Student added and Excel updated."})
    except Exception as e:return jsonify({"success":False,"message":str(e)})

@app.route("/api/edit_student",methods=["POST"])
def edit_student():
    try:
        d=request.get_json(); s=d.get("semester","Semester 1"); sid=str(d.get("Student_ID","")).strip()
        df=read_df(s)
        if df.empty:return jsonify({"success":False,"message":"Student not found."})
        df["Student_ID"]=df["Student_ID"].astype(str).str.replace(r"\.0$","",regex=True).str.strip()
        idx=df.index[df["Student_ID"]==sid].tolist()
        if not idx:return jsonify({"success":False,"message":"Student not found."})
        i=idx[0]
        for c in ["Name","Gender","Class","Attendance"]+SUBJECTS[s]: df.loc[i,c]=d.get(c,df.loc[i,c])
        save_df(df,s)
        return jsonify({"success":True,"message":"Student updated in Excel."})
    except Exception as e:return jsonify({"success":False,"message":str(e)})

@app.route("/api/delete_student",methods=["POST"])
def delete_student():
    try:
        d=request.get_json(); s=d.get("semester","Semester 1"); sid=str(d.get("Student_ID","")).strip()
        df=read_df(s); df["Student_ID"]=df["Student_ID"].astype(str).str.replace(r"\.0$","",regex=True).str.strip()
        new=df[df["Student_ID"]!=sid]
        if len(new)==len(df):return jsonify({"success":False,"message":"Student not found."})
        save_df(new,s)
        return jsonify({"success":True,"message":"Student deleted from Excel."})
    except Exception as e:return jsonify({"success":False,"message":str(e)})

if __name__=="__main__":
    app.run(host="0.0.0.0",port=int(os.environ.get("PORT",5000)))
