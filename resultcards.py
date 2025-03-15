import pandas as pd
import re
import inflect  # For converting numbers to words

p = inflect.engine()  # Initialize inflect engine

# ✅ Load Excel file
file_path = "students.xlsx"
df = pd.read_excel(file_path)

# ✅ Function to format B-Form Number
def format_bform_number(bform):
    bform = re.sub(r"[^\d]", "", str(bform))
    if len(bform) == 13:
        return f"{bform[:5]}-{bform[5:12]}-{bform[12]}"
    return bform


def clean_dob(dob):
    try:
        dob = pd.to_datetime(dob)  # Convert to datetime object
        day = dob.day
        month = dob.month
        year = dob.year

        month_name = dob.strftime("%B")  # Get full month name (e.g., "May")
        ordinal_day = f"{day}{get_ordinal_suffix(day)}"

        # Return formatted DOB in both numeric and word format
        return f"{day:02d}-{month:02d}-{year} ({ordinal_day} {month_name} {year})"
    except Exception as e:
        return str(dob)  # Return as-is if conversion fails


def get_ordinal_suffix(day):
    if 11 <= day <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")


# ✅ Function to convert DOB to words
def dob_to_words(dob):
    try:
        dob = pd.to_datetime(dob)  # Convert to datetime object
        day = p.number_to_words(dob.day).capitalize()
        month = dob.strftime("%B")  # Get full month name
        year = p.number_to_words(dob.year).capitalize()

        return f"{day} {month} {year}"
    except Exception as e:
        return str(dob)  # Return as-is if conversion fails


# ✅ Function to extract class name and section
def format_class_section(class_info):
    class_match = re.match(r"(\d+)(?:-?([A-Za-z]))?", str(class_info))
    if class_match:
        class_num = int(class_match.group(1))
        section = class_match.group(2) or ""  # Handle cases with no section
        suffix = "th" if 11 <= class_num <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(class_num % 10, "th")
        class_text = f"{class_num}{suffix}"
        return class_text, section  # Return as tuple (class_text, section)
    return class_info, ""

def get_subjects_by_class(class_num, student):
    subjects = {}

    # Define subject lists based on class (from your provided list)
    class_subjects = {
        6: ["English", "Urdu", "Mathematics", "Science", "His. Geo", "Islamiat", "Tarjama Tul Quran", "Computer", "Ethics"],
        7: ["English", "Urdu", "Mathematics", "Science", "His. Geo", "Islamiat", "Tarjama Tul Quran", "Computer", "Ethics"],
        8: ["English", "Urdu", "Mathematics", "Science", "His. Geo", "Islamiat", "Tarjama Tul Quran", "Computer", "Ethics"],
        4: ["English", "Urdu", "Mathematics", "Science", "Social Studies", "Islamiat + Nazra / Ethics"],
        5: ["English", "Urdu", "Mathematics", "Science", "Social Studies", "Islamiat + Nazra / Ethics"],
        3: ["English", "Urdu", "Mathematics", "General Knowledge", "Islamiat / Ethics"],
        2: ["English", "Urdu", "Mathematics", "General Knowledge", "Islamiat / Ethics"],
        1: ["English", "Urdu", "Mathematics", "General Knowledge", "Islamiat / Ethics"],
        "Nursery": ["English", "Urdu", "Mathematics", "General Knowledge"]
    }

    # Define subject-wise total marks (same for all classes)
    subject_marks = {
        "English": 100, "Urdu": 100, "Mathematics": 100, "Science": 100,
        "His. Geo": 100, "Islamiat": 100, "Tarjama Tul Quran": 50, "Computer": 100,
        "Ethics": 150, "Social Studies": 100, "Islamiat + Nazra / Ethics": 150,
        "General Knowledge": 100, "Islamiat / Ethics": 100
    }

    # Convert class_num to integer if it's numeric
    try:
        class_num = int(class_num)  # Convert "7" from "7-B" to integer
    except ValueError:
        pass  # If it's "Nursery", keep it as a string

    # Get subjects for this class (default to empty list if class not found)
    subject_list = class_subjects.get(class_num, [])

    # Extract only the subjects that exist in `student.xlsx`
    for subject in subject_list:
        if subject in student:  # Ensure the subject column exists in student data
            subjects[subject] = (student[subject], subject_marks[subject])
        else:
            print(f"Warning: '{subject}' not found in student data for class {class_num}")

    return subjects


# ✅ Function to assign grades
def assign_grade(percentage):
    try:
        percentage = float(percentage)
    except ValueError:
        return "Invalid"

    if percentage >= 90:
        return "A+"
    elif percentage >= 80:
        return "A"
    elif percentage >= 70:
        return "B"
    elif percentage >= 60:
        return "C"
    elif percentage >= 50:
        return "D"
    elif percentage >= 40:
        return "E"
    else:
        return "F"

# ✅ Calculate Class Position
df["Percentage"] = df["Percentage"].str.replace("%", "").astype(float)
df["Position"] = df["Percentage"].rank(method="min", ascending=False).astype(int)

# ✅ HTML Template Start
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Student Report Card</title>
    <style>
        @page {
            size: A4 landscape;
            margin: 1.2cm; /* Reduced margin slightly */
        }
        
        body {
            font-family: Arial, sans-serif;
            text-align: center;
            margin: 0;
            padding: 10px;
            background-color: #f9f9f9;
        }
        
        .page {
            display: flex;
            flex-direction: row;
            justify-content: space-between;
            align-items: start;
            width: 100%;
            height: 100vh;
            page-break-after: always;
        }
        
        .card {
            width: 48%; /* Ensures two cards fit side by side */
            height: 100%;
            box-sizing: border-box;
            padding: 20px;
            border: 2px solid black;
            page-break-inside: avoid;
            min-height: 100%;
        }
        
        .header  {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 10px;
        }
        .header-center{
            display:flex;
            flex-direction: column;
            align-items:center;
            justify-content:space-between;
        }
        .header-center h3{
                margin: 0px;
                padding: 0 10px;
                font-size: 16px;
        }
        .header img { height: 60px; }
        
        .title {
            font-size: 18px;
            font-weight: bold;
            margin:0px;
        }
        
        .info {
            text-align: left;
            font-size: 16px;
            margin: 5px 0;
            display: flex;
            gap: 20px; 
        }
        
        .table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }
        
        .table th, .table td {
            border: 1px solid black;
            padding: 5px;
            text-align: center;
            font-size: 13px;
        }
        
        .table th { background-color: #ddd; }
        
        .position-box {
            font-size: 15px;
            font-weight: bold;
            padding: 5px;
            margin-top: 10px;
            border: 2px solid black;
            text-align: center;
            width: 40%;
            margin-left: auto;
        }
        
        .top-1 { background: linear-gradient(to right, #4CAF50, #FFFF00); }
        .top-2 { background: linear-gradient(to right, #8BC34A, #FFFF80); }
        .top-3 { background: linear-gradient(to right, #CDDC39, #FFFF99); }
        
        .remarks { 
            margin-top: 10px; 
            text-align: left; 
            font-size: 14px;
        }
        
        .signature-section {
            display: flex;
            justify-content: space-between;
            margin-top: 10px;
            font-size: 14px;
        }
        
        .quote {
                padding: 5px;
                background-color: yellow;
                font-weight: bold;
                font-size: 11px;
                border-radius: 5px;
        }

        
        /* Fix print version */
        @media print {
            body { 
                margin: 0; 
                padding: 0; 
                font-size:12px;
            }
        
            .page { 
                page-break-after: always; 
                display: flex; 
                flex-wrap: wrap; 
                justify-content: space-between; 
            }
        
            .card {
                width: 45%;
                margin-bottom: 15px; 
                padding: 15px;
                page-break-inside: avoid;
                min-height: 380px; /* Prevents content from spilling */
            }
        
            .quote {
                position: absolute;
                bottom: 10px;
            }
        }


    </style>
</head>
<body>
"""
total_obtained_marks = 0  # Stores sum of obtained marks
total_full_marks = 0  # Stores sum of total marks of all subjects

# ✅ Generate HTML for Each Student (2 per A4 Page)
for index, student in df.iterrows():
    name = student["NAME"]
    father_name = student["FATHER / GUARDIAN"]
    dob = clean_dob(student["DOB"])
    bform = format_bform_number(student["FORM-B"])
    roll_number = student["ENRL #"]
    class_num = int(student["CLASS"].split('-')[0])  # Convert to integer
    subjects = get_subjects_by_class(class_num, student)
    total_obtained_marks = sum(student[subject] for subject in subjects if subject in student)
    percentage = student["Percentage"]
    grade = assign_grade(percentage)
    position = student["Position"]
    formatted_class, section = format_class_section(student["CLASS"])

    # ✅ Highlight top 3 positions
    position_class = ""
    if position == 1:
        position_class = "top-1"
    elif position == 2:
        position_class = "top-2"
    elif position == 3:
        position_class = "top-3"

    # ✅ Set total full marks based on class level
    if class_num in [6, 7, 8]:
        total_full_marks = 750
    elif class_num in [4, 5]:
        total_full_marks = 650
    elif class_num in [3, 2, 1]:
        total_full_marks = 550
    else:
        total_full_marks = 175

    # ✅ Start a new page after every two cards
    if index % 2 == 0:
        html_template += '<div class="page">'

    html_template += f"""
    <div class="card">
        <div class="header">
            <img src="gop-logo.svg" alt="Govt Logo">
            <div class="header-center">
                <h2 class="title">Report Card - Academic Year 2024-25</h2>
                <h3>Govt. High School Walton Airport Gopal Nagar Lahore</h3>
            </div>
            <img src="pec-logo.png" alt="PEC Logo" >
        </div>
        <div class="info"><span>Student Name: <u><b>{name}</b></u></span><span>Father's Name: <u><b>{father_name}</b></u></span></div>
        <div class="info"><span>B-Form: <u><b>{bform}</b></u></span><span>  Class: <u><b>{formatted_class}</b></u></span>  <span>Section: <u><b>{section}</b></u></span>  <span>Roll No: <u><b>{roll_number}</b></u></span></div>
        <div class="info"><span>Date of Birth: <u><b>{dob}</b></u> </span></div>

        <table class="table">
            <tr><th>Subject</th><th>Total Marks</th><th>Obtained Marks</th></tr>
    """

    for subject, (marks, total_marks) in subjects.items():
        html_template += f"<tr><td>{subject}</td><td>{total_marks}</td><td>{marks}</td></tr>"

    html_template += f"""
        <tr><th>Total</th><th>{total_full_marks}</th><th>{total_obtained_marks}</th></tr>
        <tr><th colspan="2">Percentage</th><td>{percentage:.2f}%</td></tr>
        <tr><th colspan="2">Grade</th><td>{grade}</td></tr>
        </table>

        <div class="position-box {position_class}">
            Class Position: {position}
        </div>

        <div class="remarks">
            Class Incharge Remarks:<br>
            ____________________________________________________________<br>
            ____________________________________________________________
        </div>

        <div class="signature-section">
            <p>Class Incharge Signature: __________</p>
            <p>Principal's Signature: ______________</p>
        </div>
        
    </div>
    """

    if index % 2 == 1 or index == len(df) - 1:
        html_template += "</div>"

html_template += "</body></html>"

with open("result_cards.html", "w", encoding="utf-8") as file:
    file.write(html_template)

print("✅ Class position added, top 3 highlighted, result cards generated!")
