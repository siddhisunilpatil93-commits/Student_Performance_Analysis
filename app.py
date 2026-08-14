from flask import Flask, render_template, jsonify, request
import pandas as pd
import os

app = Flask(__name__)

DATA_FOLDER = "data"

YEAR_FILES = {
    "1st Year": "first_year.xlsx",
    "2nd Year": "second_year.xlsx",
    "3rd Year": "third_year.xlsx"
}


def load_students(year):

    filename = YEAR_FILES.get(year)

    if not filename:
        return pd.DataFrame()

    filepath = os.path.join(DATA_FOLDER, filename)

    print("Loading file:", filepath)

    if not os.path.exists(filepath):
        print("FILE NOT FOUND:", filepath)
        return pd.DataFrame()

    try:
        df = pd.read_excel(filepath)

        df.columns = df.columns.astype(str).str.strip()

        # Convert marks to numbers
        subjects = ["OSY", "STE", "ACN", "DAN"]

        for subject in subjects:
            if subject in df.columns:
                df[subject] = pd.to_numeric(
                    df[subject],
                    errors="coerce"
                ).fillna(0)

        # Total
        existing_subjects = [
            s for s in subjects if s in df.columns
        ]

        if existing_subjects:
            df["Total"] = df[existing_subjects].sum(axis=1)

            df["Percentage"] = (
                df["Total"] /
                (len(existing_subjects) * 100)
            ) * 100

        # Attendance
        if "Attendance" in df.columns:

            df["Attendance"] = pd.to_numeric(
                df["Attendance"],
                errors="coerce"
            ).fillna(0)

            df["Attendance Status"] = df[
                "Attendance"
            ].apply(
                lambda x:
                "Good" if x >= 75 else "Low"
            )

        # Grade
        if "Percentage" in df.columns:

            def grade(p):

                if p >= 90:
                    return "A+"
                elif p >= 80:
                    return "A"
                elif p >= 70:
                    return "B"
                elif p >= 60:
                    return "C"
                elif p >= 50:
                    return "D"
                else:
                    return "F"

            df["Grade"] = df[
                "Percentage"
            ].apply(grade)

        df = df.fillna("")

        return df

    except Exception as e:

        print("ERROR READING EXCEL:", e)

        return pd.DataFrame()


# --------------------------------------------------
# HOME
# --------------------------------------------------

@app.route("/")
def home():
    return render_template("index.html")


# --------------------------------------------------
# STUDENTS
# --------------------------------------------------

@app.route("/api/students")
def students():

    year = request.args.get(
        "year",
        "1st Year"
    )

    df = load_students(year)

    return jsonify(
        df.to_dict(
            orient="records"
        )
    )


# --------------------------------------------------
# SEARCH
# --------------------------------------------------

@app.route("/api/search")
def search_student():

    year = request.args.get(
        "year",
        "1st Year"
    )

    query = request.args.get(
        "q",
        ""
    ).strip().lower()

    df = load_students(year)

    if df.empty:
        return jsonify([])

    if not query:
        return jsonify(
            df.to_dict(
                orient="records"
            )
        )

    result = df[
        df.astype(str)
        .apply(
            lambda row:
            row.str.lower()
            .str.contains(
                query,
                na=False
            )
            .any(),
            axis=1
        )
    ]

    return jsonify(
        result.to_dict(
            orient="records"
        )
    )


# --------------------------------------------------
# DASHBOARD STATS
# --------------------------------------------------

@app.route("/api/stats")
def stats():

    year = request.args.get(
        "year",
        "1st Year"
    )

    df = load_students(year)

    if df.empty:

        return jsonify({
            "total_students": 0,
            "average_percentage": 0,
            "top_performer": "-",
            "average_attendance": 0
        })

    total_students = len(df)

    # Average Percentage
    if "Percentage" in df.columns:

        percentage = pd.to_numeric(
            df["Percentage"],
            errors="coerce"
        )

        average_percentage = round(
            percentage.mean(),
            2
        )

    else:
        average_percentage = 0

    # Top Performer
    top_performer = "-"

    if (
        "Percentage" in df.columns
        and "Name" in df.columns
    ):

        temp = df.copy()

        temp["Percentage"] = pd.to_numeric(
            temp["Percentage"],
            errors="coerce"
        )

        temp = temp.dropna(
            subset=["Percentage"]
        )

        if not temp.empty:

            top_index = temp[
                "Percentage"
            ].idxmax()

            top_performer = str(
                temp.loc[
                    top_index,
                    "Name"
                ]
            )

    # Average Attendance
    if "Attendance" in df.columns:

        attendance = pd.to_numeric(
            df["Attendance"],
            errors="coerce"
        )

        average_attendance = round(
            attendance.mean(),
            2
        )

    else:
        average_attendance = 0

    return jsonify({
        "total_students": total_students,
        "average_percentage": average_percentage,
        "top_performer": top_performer,
        "average_attendance": average_attendance
    })


# --------------------------------------------------
# RUN
# --------------------------------------------------

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            5000
        )
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )
