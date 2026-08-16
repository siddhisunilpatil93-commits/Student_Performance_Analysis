import pandas as pd
import matplotlib.pyplot as plt
import os
import tkinter as tk
from tkinter import messagebox

# CREATE IMAGES FOLDER

os.makedirs("images", exist_ok=True)

# LOAD DATASET

df = pd.read_csv("data/student_data.csv")
df.columns = df.columns.str.strip()

# TOTAL MARKS

df["Total"] = (
    df["OSY"]
    + df["STE"]
    + df["ACN"]
    + df["DAN"]
)

# PERCENTAGE

df["Percentage"] = df["Total"] / 4

# GRADE

def grade(p):
    if p >= 90:
        return "A+"
    elif p >= 80:
        return "A"
    elif p >= 70:
        return "B"
    elif p >= 60:
        return "C"
    else:
        return "D"

df["Grade"] = df["Percentage"].apply(grade)

# ATTENDANCE STATUS

if "Attendance" in df.columns:
    df["Attendance Status"] = df["Attendance"].apply(
        lambda x: "Good" if x >= 75 else "Low"
    )

# ALL STUDENTS

def all_students():

    window = tk.Toplevel(login_window)

    window.title("All Students Data")
    window.geometry("1400x650")
    window.configure(bg="#EAF2F8")

    # HEADER

    header = tk.Frame(
        window,
        bg="#1F4E78"
    )
    header.pack(
        fill="x"
    )

    tk.Label(
        header,
        text="ALL STUDENTS DATA",
        font=("Arial", 20, "bold"),
        bg="#1F4E78",
        fg="white"
    ).pack(
        pady=15
    )

    # TABLE FRAME

    table_frame = tk.Frame(
        window,
        bg="#EAF2F8"
    )
    table_frame.pack(
        fill="both",
        expand=True,
        padx=15,
        pady=15
    )

    # HORIZONTAL SCROLLBAR

    x_scroll = tk.Scrollbar(
        table_frame,
        orient="horizontal"
    )
    x_scroll.pack(
        side="bottom",
        fill="x"
    )

    # VERTICAL SCROLLBAR

    y_scroll = tk.Scrollbar(
        table_frame,
        orient="vertical"
    )
    y_scroll.pack(
        side="right",
        fill="y"
    )

    # TEXT BOX

    result_text = tk.Text(
        table_frame,
        font=("Courier New", 10),
        bg="white",
        fg="#17202A",
        xscrollcommand=x_scroll.set,
        yscrollcommand=y_scroll.set,
        wrap="none"
    )

    result_text.pack(
        side="left",
        fill="both",
        expand=True
    )

    x_scroll.config(
        command=result_text.xview
    )

    y_scroll.config(
        command=result_text.yview
    )

    # HEADER ROW

    header_text = (
        f"{'Student_ID':<12}"
        f"{'Name':<18}"
        f"{'Gender':<10}"
        f"{'Class':<10}"
        f"{'OSY':<8}"
        f"{'STE':<8}"
        f"{'ACN':<8}"
        f"{'DAN':<8}"
        f"{'Total':<10}"
        f"{'Percentage':<13}"
        f"{'Attendance':<12}"
        f"{'Attendance Status':<20}"
        f"{'Grade':<8}\n"
    )

    result_text.insert(
        tk.END,
        header_text
    )

    result_text.insert(
        tk.END,
        "=" * 170 + "\n"
    )

    # DISPLAY ALL STUDENTS

    for index, student in df.iterrows():

        row = (
            f"{str(student['Student_ID']):<12}"
            f"{str(student['Name']):<18}"
            f"{str(student['Gender']):<10}"
            f"{str(student['Class']):<10}"
            f"{str(student['OSY']):<8}"
            f"{str(student['STE']):<8}"
            f"{str(student['ACN']):<8}"
            f"{str(student['DAN']):<8}"
            f"{str(student['Total']):<10}"
            f"{student['Percentage']:.2f}%"
            f"{'':<7}"
            f"{str(student['Attendance']) + '%':<12}"
            f"{str(student['Attendance Status']):<20}"
            f"{str(student['Grade']):<8}\n"
        )

        result_text.insert(
            tk.END,
            row
        )

    result_text.config(
        state="disabled"
    )


    
    

# STUDENT DETAILS

def student_details():

    window = tk.Toplevel(login_window)

    window.title("Student Details")
    window.geometry("750x600")
    window.configure(
        bg="#EAF2F8"
    )

    # HEADER

    tk.Label(
        window,
        text="STUDENT DETAILS",
        font=("Arial", 20, "bold"),
        bg="#1F4E78",
        fg="white"
    ).pack(
        fill="x",
        pady=10
    )

    # SEARCH LABEL

    tk.Label(
        window,
        text="Enter Student Name or Student ID:",
        font=("Arial", 12, "bold"),
        bg="#EAF2F8"
    ).pack(
        pady=10
    )

    # SEARCH ENTRY

    search_entry = tk.Entry(
        window,
        width=35,
        font=("Arial", 12)
    )

    search_entry.pack(
        pady=5
    )

    # RESULT TEXT

    result_text = tk.Text(
        window,
        width=85,
        height=22,
        font=("Courier New", 10)
    )

    result_text.pack(
        pady=15
    )

    # SEARCH FUNCTION

    def search_student():

        search_value = search_entry.get().strip()

        result_text.delete(
            "1.0",
            tk.END
        )

        if search_value == "":
            result_text.insert(
                tk.END,
                "Please enter Student Name or Student ID."
            )
            return

        # SEARCH BY NAME PART OR STUDENT ID

        result = df[
            (
                df["Name"]
                .astype(str)
                .str.lower()
                .str.contains(
                    search_value.lower(),
                    na=False
                )
            )
            |
            (
                df["Student_ID"]
                .astype(str)
                .str.strip()
                == search_value
            )
        ]

        # STUDENT NOT FOUND

        if result.empty:

            result_text.insert(
                tk.END,
                "Student Not Found"
            )

        else:

            # DISPLAY ALL MATCHING STUDENTS

            for index, student in result.iterrows():

                result_text.insert(
                    tk.END,
                    "========== STUDENT DETAILS ==========\n\n"
                )

                result_text.insert(
                    tk.END,
                    f"Student ID        : {student['Student_ID']}\n"
                )

                result_text.insert(
                    tk.END,
                    f"Name              : {student['Name']}\n"
                )

                result_text.insert(
                    tk.END,
                    f"Gender            : {student['Gender']}\n"
                )

                result_text.insert(
                    tk.END,
                    f"Class             : {student['Class']}\n"
                )

                result_text.insert(
                    tk.END,
                    f"OSY Marks         : {student['OSY']}\n"
                )

                result_text.insert(
                    tk.END,
                    f"STE Marks         : {student['STE']}\n"
                )

                result_text.insert(
                    tk.END,
                    f"ACN Marks         : {student['ACN']}\n"
                )

                result_text.insert(
                    tk.END,
                    f"DAN Marks         : {student['DAN']}\n"
                )

                result_text.insert(
                    tk.END,
                    f"Total Marks       : {student['Total']}\n"
                )

                result_text.insert(
                    tk.END,
                    f"Percentage        : {student['Percentage']:.2f}%\n"
                )

                result_text.insert(
                    tk.END,
                    f"Attendance        : {student['Attendance']}%\n"
                )

                result_text.insert(
                    tk.END,
                    f"Attendance Status : {student['Attendance Status']}\n"
                )

                result_text.insert(
                    tk.END,
                    f"Grade             : {student['Grade']}\n"
                )

                result_text.insert(
                    tk.END,
                    "\n" + "-" * 55 + "\n\n"
                )

    # SEARCH BUTTON

    tk.Button(
        window,
        text="SEARCH",
        width=15,
        height=2,
        font=("Arial", 11, "bold"),
        bg="#8E44AD",
        fg="white",
        activebackground="#6C3483",
        activeforeground="white",
        relief="flat",
        cursor="hand2",
        command=search_student
    ).pack(
        pady=5
    )


# TOP 5 STUDENTS

def top_5_students():

    top_5 = df.sort_values(
        by="Percentage",
        ascending=False
    ).head(5)

    window = tk.Toplevel(login_window)

    window.title("Top 5 Students")
    window.geometry("800x500")

    window.configure(
        bg="#EAF2F8"
    )

    tk.Label(
        window,
        text="TOP 5 STUDENTS",
        font=("Arial", 20, "bold"),
        bg="#F1C40F",
        fg="black"
    ).pack(
        fill="x",
        pady=10
    )

    result_text = tk.Text(
        window,
        width=90,
        height=20,
        font=("Courier New", 11)
    )

    result_text.pack(
        pady=20
    )

    for index, student in top_5.iterrows():

        result_text.insert(
            tk.END,
            f"Name        : {student['Name']}\n"
        )

        result_text.insert(
            tk.END,
            f"Total Marks : {student['Total']}\n"
        )

        result_text.insert(
            tk.END,
            f"Percentage  : {student['Percentage']:.2f}%\n"
        )

        result_text.insert(
            tk.END,
            f"Grade       : {student['Grade']}\n"
        )

        result_text.insert(
            tk.END,
            "-" * 50 + "\n\n"
        )


# AVERAGE MARKS

def average_marks():

    subjects = [
        "OSY",
        "STE",
        "ACN",
        "DAN"
    ]

    average = df[subjects].mean()

    plt.figure(
        figsize=(7, 5)
    )

    plt.bar(
        subjects,
        average
    )

    plt.title(
        "Average Marks"
    )

    plt.xlabel(
        "Subjects"
    )

    plt.ylabel(
        "Average Marks"
    )

    plt.grid(
        axis="y"
    )

    plt.show()


# GRADE DISTRIBUTION

def grade_distribution():

    grade_count = df["Grade"].value_counts()

    plt.figure(
        figsize=(6, 6)
    )

    plt.pie(
        grade_count,
        labels=grade_count.index,
        autopct="%1.1f%%",
        startangle=90
    )

    plt.title(
        "Grade Distribution"
    )

    plt.show()


# LOGIN

def login():

    username = username_entry.get()
    password = password_entry.get()

    if username == "silicon" and password == "patil":

        messagebox.showinfo(
            "Login Successful",
            "Welcome to Student Performance System!"
        )

        login_window.withdraw()

        dashboard()

    else:

        messagebox.showerror(
            "Login Failed",
            "Invalid Username or Password"
        )


# COLORFUL DASHBOARD

def dashboard():

    dashboard_window = tk.Toplevel(
        login_window
    )

    dashboard_window.title(
        "Student Performance Dashboard"
    )

    dashboard_window.geometry(
        "950x650"
    )

    dashboard_window.configure(
        bg="#EAF2F8"
    )

    # HEADER

    header = tk.Frame(
        dashboard_window,
        bg="#1F4E78",
        height=100
    )

    header.pack(
        fill="x"
    )

    tk.Label(
        header,
        text="STUDENT PERFORMANCE DASHBOARD",
        font=("Arial", 24, "bold"),
        bg="#1F4E78",
        fg="white"
    ).pack(
        pady=(20, 5)
    )

    tk.Label(
        header,
        text="Department of Computer Engineering",
        font=("Arial", 11),
        bg="#1F4E78",
        fg="white"
    ).pack()

    # INFORMATION CARDS

    info_frame = tk.Frame(
        dashboard_window,
        bg="#EAF2F8"
    )

    info_frame.pack(
        pady=25
    )

    # TOTAL STUDENTS CARD

    total_students = len(df)

    total_card = tk.Frame(
        info_frame,
        bg="#3498DB",
        width=220,
        height=120,
        cursor="hand2"
    )

    total_card.grid(
        row=0,
        column=0,
        padx=15
    )

    total_card.pack_propagate(
        False
    )

    total_title = tk.Label(
        total_card,
        text="TOTAL STUDENTS",
        font=("Arial", 12, "bold"),
        bg="#3498DB",
        fg="white",
        cursor="hand2"
    )

    total_title.pack(
        pady=(20, 5)
    )

    total_number = tk.Label(
        total_card,
        text=str(total_students),
        font=("Arial", 28, "bold"),
        bg="#3498DB",
        fg="white",
        cursor="hand2"
    )

    total_number.pack()

    # CLICK TOTAL STUDENTS

    total_card.bind(
        "<Button-1>",
        lambda event: all_students()
    )

    total_title.bind(
        "<Button-1>",
        lambda event: all_students()
    )

    total_number.bind(
        "<Button-1>",
        lambda event: all_students()
    )

    # AVERAGE PERCENTAGE CARD

    avg_percentage = df["Percentage"].mean()

    avg_card = tk.Frame(
        info_frame,
        bg="#27AE60",
        width=220,
        height=120
    )

    avg_card.grid(
        row=0,
        column=1,
        padx=15
    )

    avg_card.pack_propagate(
        False
    )

    tk.Label(
        avg_card,
        text="AVERAGE PERCENTAGE",
        font=("Arial", 12, "bold"),
        bg="#27AE60",
        fg="white"
    ).pack(
        pady=(20, 5)
    )

    tk.Label(
        avg_card,
        text=f"{avg_percentage:.2f}%",
        font=("Arial", 28, "bold"),
        bg="#27AE60",
        fg="white"
    ).pack()

    # TOP STUDENT CARD

    top_student = df.loc[
        df["Percentage"].idxmax(),
        "Name"
    ]

    top_percentage = df["Percentage"].max()

    top_card = tk.Frame(
        info_frame,
        bg="#E67E22",
        width=220,
        height=120
    )

    top_card.grid(
        row=0,
        column=2,
        padx=15
    )

    top_card.pack_propagate(
        False
    )

    tk.Label(
        top_card,
        text="TOP STUDENT",
        font=("Arial", 12, "bold"),
        bg="#E67E22",
        fg="white"
    ).pack(
        pady=(15, 3)
    )

    tk.Label(
        top_card,
        text=top_student,
        font=("Arial", 16, "bold"),
        bg="#E67E22",
        fg="white"
    ).pack()

    tk.Label(
        top_card,
        text=f"{top_percentage:.2f}%",
        font=("Arial", 13, "bold"),
        bg="#E67E22",
        fg="white"
    ).pack()

    # BUTTON FRAME

    button_frame = tk.Frame(
        dashboard_window,
        bg="#EAF2F8"
    )

    button_frame.pack(
        pady=10
    )

    # STUDENT DETAILS

    tk.Button(
        button_frame,
        text="Student Details",
        width=28,
        height=2,
        font=("Arial", 12, "bold"),
        bg="#8E44AD",
        fg="white",
        activebackground="#6C3483",
        relief="flat",
        cursor="hand2",
        command=student_details
    ).grid(
        row=0,
        column=0,
        padx=15,
        pady=10
    )

    # TOP 5

    tk.Button(
        button_frame,
        text="Top 5 Students",
        width=28,
        height=2,
        font=("Arial", 12, "bold"),
        bg="#F1C40F",
        fg="black",
        activebackground="#D4AC0D",
        relief="flat",
        cursor="hand2",
        command=top_5_students
    ).grid(
        row=0,
        column=1,
        padx=15,
        pady=10
    )

    # AVERAGE MARKS

    tk.Button(
        button_frame,
        text="Average Marks",
        width=28,
        height=2,
        font=("Arial", 12, "bold"),
        bg="#16A085",
        fg="white",
        activebackground="#117864",
        relief="flat",
        cursor="hand2",
        command=average_marks
    ).grid(
        row=1,
        column=0,
        padx=15,
        pady=10
    )

    # GRADE DISTRIBUTION

    tk.Button(
        button_frame,
        text="Grade Distribution",
        width=28,
        height=2,
        font=("Arial", 12, "bold"),
        bg="#E74C3C",
        fg="white",
        activebackground="#C0392B",
        relief="flat",
        cursor="hand2",
        command=grade_distribution
    ).grid(
        row=1,
        column=1,
        padx=15,
        pady=10
    )

    # EXIT

    tk.Button(
        dashboard_window,
        text="EXIT",
        width=20,
        height=2,
        font=("Arial", 11, "bold"),
        bg="#34495E",
        fg="white",
        activebackground="#2C3E50",
        relief="flat",
        cursor="hand2",
        command=dashboard_window.destroy
    ).pack(
        pady=20
    )


# LOGIN WINDOW

login_window = tk.Tk()

login_window.title(
    "Student Performance - Login"
)

login_window.geometry(
    "500x350"
)

login_window.configure(
    bg="#EAF2F8"
)

# LOGIN HEADER

tk.Label(
    login_window,
    text="STUDENT PERFORMANCE SYSTEM",
    font=("Arial", 20, "bold"),
    bg="#1F4E78",
    fg="white"
).pack(
    fill="x",
    pady=20
)

tk.Label(
    login_window,
    text="DEPARTMENT OF COMPUTER ENGINEERING",
    font=("Arial", 12, "bold"),
    bg="#EAF2F8"
).pack(
    pady=5
)

# USERNAME

tk.Label(
    login_window,
    text="Username",
    font=("Arial", 11, "bold"),
    bg="#EAF2F8"
).pack(
    pady=10
)

username_entry = tk.Entry(
    login_window,
    width=30
)

username_entry.pack()

# PASSWORD

tk.Label(
    login_window,
    text="Password",
    font=("Arial", 11, "bold"),
    bg="#EAF2F8"
).pack(
    pady=10
)

password_entry = tk.Entry(
    login_window,
    width=30,
    show="*"
)

password_entry.pack()

# LOGIN BUTTON

tk.Button(
    login_window,
    text="LOGIN",
    width=15,
    height=2,
    font=("Arial", 11, "bold"),
    bg="#27AE60",
    fg="white",
    command=login
).pack(
    pady=25
)

# START

login_window.mainloop()