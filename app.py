from flask import Flask, request, send_file, render_template
import pandas as pd
import os

app = Flask(__name__)

def read_student_preferences(file_path):
    df = pd.read_excel(file_path)
    student_preferences = []
    for index, row in df.iterrows():
        student = row[0]
        preferences = row[1:].dropna().tolist()
        student_preferences.append({'student': student, 'preferences': preferences})
    return student_preferences

def read_faculty_preferences(file_path):
    df = pd.read_excel(file_path)
    faculty_order = df.iloc[:, 0].tolist()
    faculty_preferences = {}
    for index, row in df.iterrows():
        faculty = row[0]
        preferences = row[1:].dropna().tolist()
        faculty_preferences[faculty] = preferences
    return faculty_preferences, faculty_order

def allocate_students(student_preferences, faculty_preferences):
    allocated_students = {faculty: [] for faculty in faculty_preferences.keys()}
    all_students = student_preferences[:]
    preference_level = 0
    max_students_per_faculty = 8
    while all_students and preference_level < 5:
        students_left = []
        for student_data in all_students:
            student = student_data['student']
            if preference_level < len(student_data['preferences']):
                preferred_faculty = student_data['preferences'][preference_level]
                if preferred_faculty in faculty_preferences:
                    if len(faculty_preferences[preferred_faculty]) > 0 and len(allocated_students[preferred_faculty]) < max_students_per_faculty:
                        faculty_first_preference = faculty_preferences[preferred_faculty][0]
                        if student == faculty_first_preference:
                            allocated_students[preferred_faculty].append(student)
                            for faculty in faculty_preferences:
                                faculty_preferences[faculty] = [x for x in faculty_preferences[faculty] if x != student]
                            continue
            students_left.append(student_data)
        all_students = students_left
        preference_level += 1
    return allocated_students

def save_allocations_to_excel(allocations, faculty_order):
    allocations_data = {'Faculty': faculty_order, 'Students': []}
    for faculty in faculty_order:
        students = allocations.get(faculty, [])
        allocations_data['Students'].append(', '.join(students))
    allocations_df = pd.DataFrame(allocations_data)
    allocations_file = 'allocations.xlsx'
    allocations_df.to_excel(allocations_file, index=False)
    return allocations_file

def save_unallocated_students_to_excel(unallocated_students):
    unallocated_df = pd.DataFrame({'Unallocated Students': unallocated_students})
    unallocated_file = 'unallocated_students.xlsx'
    unallocated_df.to_excel(unallocated_file, index=False)
    return unallocated_file

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload_student', methods=['POST'])
def upload_student():
    student_file = request.files['student_file']
    student_file_path = 'uploaded_student.xlsx'
    student_file.save(student_file_path)
    return 'Student file uploaded successfully.'

@app.route('/upload_faculty', methods=['POST'])
def upload_faculty():
    faculty_file = request.files['faculty_file']
    faculty_file_path = 'uploaded_faculty.xlsx'
    faculty_file.save(faculty_file_path)
    return 'Faculty file uploaded successfully.'

@app.route('/allocate', methods=['POST'])
def allocate():
    student_file_path = 'uploaded_student.xlsx'
    faculty_file_path = 'uploaded_faculty.xlsx'

    student_preferences = read_student_preferences(student_file_path)
    faculty_preferences, faculty_order = read_faculty_preferences(faculty_file_path)

    allocated_students = allocate_students(student_preferences, faculty_preferences)

    all_students_set = {student['student'] for student in student_preferences}
    allocated_students_set = set(sum(allocated_students.values(), []))
    unallocated_students = list(all_students_set - allocated_students_set)

    allocations_file = save_allocations_to_excel(allocated_students, faculty_order)
    unallocated_file = save_unallocated_students_to_excel(unallocated_students)

    return {
        'allocations': allocations_file,
        'unallocated': unallocated_file
    }

@app.route('/download/<filename>')
def download(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
