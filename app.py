from flask import Flask, render_template, jsonify, request, redirect, url_for, session
import pandas as pd
import os

app = Flask(__name__)

app.secret_key = os.environ.get(
    'SECRET_KEY',
    'student-performance-secret-key'
)

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))


# =========================================================
# SEMESTER -> YEAR
# =========================================================

SEMESTER_YEAR = {
    'Semester 1': '1st Year',
    'Semester 2': '1st Year',
    'Semester 3': '2nd Year',
    'Semester 4': '2nd Year',
    'Semester 5': '3rd Year',
    'Semester 6': '3rd Year'
}


# =========================================================
# EXCEL FILES
# =========================================================

SEMESTER_FILES = {
    f'Semester {i}': f'semester_{i}.xlsx'
    for i in range(1, 7)
}


# =========================================================
# SUBJECTS
# =========================================================

SUBJECTS = {

    'Semester 1': [
        'Basic Mathematics',
        'Communication Skills',
        'Basic Science'
    ],

    'Semester 2': [
        'Applied Mathematics',
        'Basic Electrical and Engineering',
        'Programming in C'
    ],

    'Semester 3': [
        'Object Oriented Programming',
        'Data Structure',
        'Digital Techniques',
        'Database Management System'
    ],

    'Semester 4': [
        'Java Programming',
        'Data Communication and Network',
        'Microprocessor Programming',
        'Environmental Education And Sustanability'
    ],

    'Semester 5': [
        'Software Engineering',
        'Opreting System',
        'Data Analytics'
    ],

    'Semester 6': [
        'Mobile Application Development',
        'Machine Learning',
        'Software Testing',
        'Mangement'
    ]
}


# =========================================================
# LOGIN
# =========================================================

LOGIN_USERNAME = os.environ.get(
    'ADMIN_USERNAME',
    'silicon'
)

LOGIN_PASSWORD = os.environ.get(
    'ADMIN_PASSWORD',
    'patil'
)


# =========================================================
# HELPERS
# =========================================================

def valid_semester(s):
    return s in SEMESTER_FILES


def excel_path(s):
    return os.path.join(
        BASE_FOLDER,
        SEMESTER_FILES[s]
    )


def logged_in():
    return session.get('logged_in') is True


def page_login_check():

    if not logged_in():
        return redirect(url_for('login'))

    return None


def api_login_check():

    if not logged_in():

        return jsonify({
            'success': False,
            'message': 'Login required.'
        }), 401

    return None


# =========================================================
# GRADE
# =========================================================

def calculate_grade(p):

    try:
        p = float(p)
    except Exception:
        return 'F'

    if p >= 90:
        return 'A+'
    elif p >= 80:
        return 'A'
    elif p >= 70:
        return 'B+'
    elif p >= 60:
        return 'B'
    elif p >= 50:
        return 'C'
    elif p >= 40:
        return 'D'
    else:
        return 'F'


# =========================================================
# READ EXCEL
# =========================================================

def read_excel(s):

    if not valid_semester(s):
        return pd.DataFrame()

    path = excel_path(s)

    if not os.path.exists(path):
        return pd.DataFrame()

    try:

        df = pd.read_excel(path)

        if df is None:
            return pd.DataFrame()

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        return df

    except Exception as e:

        print('Excel read error:', e)
        return pd.DataFrame()


# =========================================================
# PROCESS DATA
# =========================================================

def process_data(df, s):

    df = df.copy()

    required = [
        'Student_ID',
        'Name',
        'Gender',
        'Class',
        'Attendance'
    ]

    for c in required:

        if c not in df.columns:

            if c == 'Attendance':

                df[c] = pd.Series(
                    0,
                    index=df.index,
                    dtype=float
                )

            else:

                df[c] = pd.Series(
                    '',
                    index=df.index,
                    dtype=str
                )


    # Student ID

    df['Student_ID'] = (
        df['Student_ID']
        .fillna('')
        .astype(str)
        .str.replace(
            r'\.0$',
            '',
            regex=True
        )
        .str.strip()
    )


    # Text

    for c in [
        'Name',
        'Gender',
        'Class'
    ]:

        df[c] = (
            df[c]
            .fillna('')
            .astype(str)
            .str.strip()
        )


    # Attendance

    raw_attendance = (
        df['Attendance']
        .astype(str)
        .str.strip()
    )

    percent_mask = (
        raw_attendance.str.endswith('%')
    )

    clean_attendance = (
        raw_attendance
        .str.replace(
            '%',
            '',
            regex=False
        )
    )

    df['Attendance'] = pd.to_numeric(
        clean_attendance,
        errors='coerce'
    )

    fraction_mask = (
        df['Attendance'].notna()
        &
        df['Attendance'].between(0, 1)
        &
        ~percent_mask
    )

    df.loc[
        fraction_mask,
        'Attendance'
    ] = (
        df.loc[
            fraction_mask,
            'Attendance'
        ] * 100
    )

    df['Attendance'] = (
        df['Attendance']
        .fillna(0)
        .clip(0, 100)
        .round(2)
    )


    # Subjects

    subs = SUBJECTS.get(s, [])

    for c in subs:

        if c in df.columns:

            df[c] = pd.to_numeric(
                df[c],
                errors='coerce'
            )

            df[c] = (
                df[c]
                .fillna(0)
                .clip(0, 100)
                .round(2)
            )

        else:

            df[c] = pd.Series(
                0,
                index=df.index,
                dtype=float
            )


    # Total / Percentage

    if subs:

        df['Total'] = (
            df[subs]
            .sum(axis=1)
            .round(2)
        )

        df['Percentage'] = (
            df['Total']
            / (len(subs) * 100)
            * 100
        ).round(2)

    else:

        df['Total'] = 0.0
        df['Percentage'] = 0.0


    # Attendance status

    df['Attendance Status'] = (
        df['Attendance']
        .apply(
            lambda x:
            'Good'
            if float(x) >= 75
            else 'Bad'
        )
    )


    # Grade

    df['Grade'] = (
        df['Percentage']
        .apply(calculate_grade)
    )

    return df


# =========================================================
# SAVE EXCEL
# =========================================================

def save_excel(df, s):

    processed = process_data(
        df,
        s
    )

    cols = (
        [
            'Student_ID',
            'Name',
            'Gender',
            'Class'
        ]
        + SUBJECTS[s]
        + [
            'Attendance'
        ]
    )

    for c in cols:

        if c not in processed.columns:

            if c == 'Attendance':

                processed[c] = pd.Series(
                    0,
                    index=processed.index,
                    dtype=float
                )

            else:

                processed[c] = pd.Series(
                    '',
                    index=processed.index,
                    dtype=str
                )

    processed[cols].to_excel(
        excel_path(s),
        index=False
    )


# =========================================================
# LOGIN
# =========================================================

@app.route(
    '/login',
    methods=['GET', 'POST']
)
def login():

    if logged_in():
        return redirect(url_for('home'))

    if request.method == 'POST':

        username = request.form.get(
            'username',
            ''
        ).strip()

        password = request.form.get(
            'password',
            ''
        )

        if (
            username == LOGIN_USERNAME
            and password == LOGIN_PASSWORD
        ):

            session.clear()
            session['logged_in'] = True

            return redirect(
                url_for('home')
            )

        return render_template(
            'index.html',
            login_page=True,
            login_error='Invalid username or password.'
        )

    return render_template(
        'index.html',
        login_page=True
    )


# =========================================================
# LOGOUT
# =========================================================

@app.route('/logout')
def logout():

    session.clear()

    return redirect(
        url_for('login')
    )


# =========================================================
# HOME
# =========================================================

@app.route('/')
def home():

    check = page_login_check()

    if check:
        return check

    return render_template(
        'index.html',
        login_page=False
    )


# =========================================================
# SUBJECT API
# =========================================================

@app.route('/api/subjects')
def api_subjects():

    check = api_login_check()

    if check:
        return check

    s = request.args.get(
        'semester',
        'Semester 1'
    ).strip()

    if not valid_semester(s):

        return jsonify({
            'success': False,
            'message': 'Invalid semester'
        }), 400

    return jsonify({

        'success': True,

        'branch':
            'Computer Engineering',

        'year':
            SEMESTER_YEAR[s],

        'semester':
            s,

        'subjects': [
            {
                'code': x,
                'name': x
            }
            for x in SUBJECTS[s]
        ]
    })


# =========================================================
# STUDENTS API
# =========================================================

@app.route('/api/students')
def api_students():

    check = api_login_check()

    if check:
        return check

    s = request.args.get(
        'semester',
        'Semester 1'
    ).strip()

    if not valid_semester(s):
        return jsonify([])

    df = read_excel(s)

    if df.empty:
        return jsonify([])

    df = process_data(
        df,
        s
    )

    cols = (
        [
            'Student_ID',
            'Name',
            'Gender',
            'Class'
        ]
        + SUBJECTS[s]
        + [
            'Total',
            'Percentage',
            'Attendance',
            'Attendance Status',
            'Grade'
        ]
    )

    return jsonify(
        df[cols]
        .fillna('')
        .to_dict('records')
    )


# =========================================================
# SINGLE SEMESTER ANALYTICS
# =========================================================

@app.route('/api/analytics')
def api_analytics():

    check = api_login_check()

    if check:
        return check

    s = request.args.get(
        'semester',
        'Semester 1'
    ).strip()

    empty = {

        'total_students': 0,

        'average_percentage': 0,

        'top_performer': '-',

        'average_attendance': 0,

        'subjects': {},

        'grades': {}
    }

    if not valid_semester(s):
        return jsonify(empty)

    df = read_excel(s)

    if df.empty:
        return jsonify(empty)

    df = process_data(
        df,
        s
    )

    if df.empty:
        return jsonify(empty)

    try:

        top_index = df[
            'Percentage'
        ].idxmax()

        top = df.loc[
            top_index,
            'Name'
        ]

    except Exception:

        top = '-'

    return jsonify({

        'total_students':
            int(len(df)),

        'average_percentage':
            round(
                float(
                    df['Percentage'].mean()
                ),
                2
            ),

        'top_performer':
            str(top),

        'average_attendance':
            round(
                float(
                    df['Attendance'].mean()
                ),
                2
            ),

        'subjects': {
            x:
            round(
                float(
                    df[x].mean()
                ),
                2
            )
            for x in SUBJECTS[s]
        },

        'grades':
            df[
                'Grade'
            ].value_counts().to_dict()
    })


# =========================================================
# COMBINED SEMESTER 1 + 2 ANALYTICS
# =========================================================

@app.route('/api/combined_analytics')
def combined_analytics():

    check = api_login_check()

    if check:
        return check

    semesters = [
        'Semester 1',
        'Semester 2'
    ]

    result = {}

    for s in semesters:

        df = read_excel(s)

        result[s] = {
            'subjects': {},
            'grades': {},
            'students': 0,
            'percentage': 0,
            'attendance': 0
        }

        if df.empty:
            continue

        df = process_data(
            df,
            s
        )

        if df.empty:
            continue

        result[s]['students'] = int(
            len(df)
        )

        result[s]['percentage'] = round(
            float(
                df['Percentage'].mean()
            ),
            2
        )

        result[s]['attendance'] = round(
            float(
                df['Attendance'].mean()
            ),
            2
        )

        result[s]['subjects'] = {

            subject:
            round(
                float(
                    df[subject].mean()
                ),
                2
            )

            for subject in SUBJECTS[s]
        }

        result[s]['grades'] = (
            df['Grade']
            .value_counts()
            .to_dict()
        )

    return jsonify({
        'success': True,
        'semesters': result
    })


# =========================================================
# UPLOAD EXCEL
# =========================================================

@app.route(
    '/api/upload_excel',
    methods=['POST']
)
def upload_excel():

    check = api_login_check()

    if check:
        return check

    try:

        s = request.form.get(
            'semester',
            ''
        ).strip()

        f = request.files.get(
            'file'
        )

        if not valid_semester(s):

            return jsonify({
                'success': False,
                'message':
                    'Invalid semester selected.'
            }), 400

        if not f or not f.filename:

            return jsonify({
                'success': False,
                'message':
                    'Please select Excel file.'
            }), 400

        if not f.filename.lower().endswith(
            ('.xlsx', '.xls')
        ):

            return jsonify({
                'success': False,
                'message':
                    'Only Excel files are allowed.'
            }), 400

        df = pd.read_excel(f)

        if df.empty:

            return jsonify({
                'success': False,
                'message':
                    'Uploaded Excel is empty.'
            }), 400

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        for c in [
            'Student_ID',
            'Name'
        ]:

            if c not in df.columns:

                return jsonify({
                    'success': False,
                    'message':
                        f'{c} column is required.'
                }), 400

        save_excel(
            df,
            s
        )

        return jsonify({

            'success': True,

            'message':
                f'{s} Excel uploaded successfully.'
        })

    except Exception as e:

        print(
            'Excel upload error:',
            repr(e)
        )

        return jsonify({

            'success': False,

            'message':
                f'Excel upload failed: {e}'
        }), 500


# =========================================================
# ADD STUDENT
# =========================================================

@app.route(
    '/api/add_student',
    methods=['POST']
)
def add_student():

    check = api_login_check()

    if check:
        return check

    try:

        d = request.get_json() or {}

        s = str(
            d.get(
                'semester',
                'Semester 1'
            )
        ).strip()

        sid = str(
            d.get(
                'Student_ID',
                ''
            )
        ).strip()

        name = str(
            d.get(
                'Name',
                ''
            )
        ).strip()

        if not valid_semester(s):

            return jsonify({
                'success': False,
                'message':
                    'Invalid semester.'
            }), 400

        if not sid or not name:

            return jsonify({
                'success': False,
                'message':
                    'Student ID and Name are required.'
            })

        df = read_excel(s)

        if (
            not df.empty
            and 'Student_ID' in df.columns
        ):

            ids = (
                df['Student_ID']
                .fillna('')
                .astype(str)
                .str.replace(
                    r'\.0$',
                    '',
                    regex=True
                )
                .str.strip()
            )

            if sid in ids.tolist():

                return jsonify({
                    'success': False,
                    'message':
                        'Student ID already exists.'
                })

        row = {

            'Student_ID': sid,

            'Name': name,

            'Gender': d.get(
                'Gender',
                ''
            ),

            'Class': d.get(
                'Class',
                ''
            ),

            'Attendance': d.get(
                'Attendance',
                0
            )
        }

        row.update({

            x: d.get(
                x,
                0
            )

            for x in SUBJECTS[s]
        })

        df = pd.concat(
            [
                df,
                pd.DataFrame([row])
            ],
            ignore_index=True
        )

        save_excel(
            df,
            s
        )

        return jsonify({

            'success': True,

            'message':
                'Student added successfully.'
        })

    except Exception as e:

        return jsonify({

            'success': False,

            'message':
                str(e)
        }), 500


# =========================================================
# EDIT STUDENT
# =========================================================

@app.route(
    '/api/edit_student',
    methods=['POST']
)
def edit_student():

    check = api_login_check()

    if check:
        return check

    try:

        d = request.get_json() or {}

        s = str(
            d.get(
                'semester',
                ''
            )
        ).strip()

        sid = str(
            d.get(
                'Student_ID',
                ''
            )
        ).strip()

        if not valid_semester(s):

            return jsonify({
                'success': False,
                'message':
                    'Invalid semester.'
            }), 400

        df = read_excel(s)

        if (
            df.empty
            or 'Student_ID' not in df.columns
        ):

            return jsonify({
                'success': False,
                'message':
                    'Student not found.'
            })

        df['Student_ID'] = (
            df['Student_ID']
            .fillna('')
            .astype(str)
            .str.replace(
                r'\.0$',
                '',
                regex=True
            )
            .str.strip()
        )

        matches = df.index[
            df['Student_ID'] == sid
        ].tolist()

        if not matches:

            return jsonify({
                'success': False,
                'message':
                    'Student not found.'
            })

        i = matches[0]

        for c in (
            [
                'Name',
                'Gender',
                'Class',
                'Attendance'
            ]
            + SUBJECTS[s]
        ):

            if c in d:

                df.loc[
                    i,
                    c
                ] = d[c]

        save_excel(
            df,
            s
        )

        return jsonify({

            'success': True,

            'message':
                'Student updated successfully.'
        })

    except Exception as e:

        return jsonify({

            'success': False,

            'message':
                str(e)
        }), 500


# =========================================================
# DELETE STUDENT
# =========================================================

@app.route(
    '/api/delete_student',
    methods=['POST']
)
def delete_student():

    check = api_login_check()

    if check:
        return check

    try:

        d = request.get_json() or {}

        s = str(
            d.get(
                'semester',
                ''
            )
        ).strip()

        sid = str(
            d.get(
                'Student_ID',
                ''
            )
        ).strip()

        if not valid_semester(s):

            return jsonify({
                'success': False,
                'message':
                    'Invalid semester.'
            }), 400

        df = read_excel(s)

        if df.empty:

            return jsonify({
                'success': False,
                'message':
                    'Student not found.'
            })

        df['Student_ID'] = (
            df['Student_ID']
            .fillna('')
            .astype(str)
            .str.replace(
                r'\.0$',
                '',
                regex=True
            )
            .str.strip()
        )

        new_df = df[
            df['Student_ID'] != sid
        ].copy()

        if len(new_df) == len(df):

            return jsonify({
                'success': False,
                'message':
                    'Student not found.'
            })

        save_excel(
            new_df,
            s
        )

        return jsonify({

            'success': True,

            'message':
                'Student deleted successfully.'
        })

    except Exception as e:

        return jsonify({

            'success': False,

            'message':
                str(e)
        }), 500


# =========================================================
# HEALTH
# =========================================================

@app.route('/health')
def health():

    return jsonify({

        'status': 'ok',

        'application':
            'Student Performance Analysis System'
    })


# =========================================================
# RUN
# =========================================================

if __name__ == '__main__':

    app.run(
        host='0.0.0.0',
        port=int(
            os.environ.get(
                'PORT',
                5000
            )
        )
    )
