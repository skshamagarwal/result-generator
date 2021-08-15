from datetime import date
from fpdf import FPDF
from pandas import read_excel
from datetime import date

df = read_excel('Dummy Data.xlsx', sheet_name='Sheet1')

# Arranging Dummy Data in usable format
records = []
scorecard = []
counter = 1
noOfQs = 25
for i in range(1, len(df)):

    # Personal Records
    candNo = df.loc[i][0]
    roundN = df.loc[i][1]
    firstName = df.loc[i][2]
    lastName = df.loc[i][3]
    fullName = df.loc[i][4]
    regNo = df.loc[i][5]
    grade = df.loc[i][6]
    school = df.loc[i][7]
    gender = df.loc[i][8]
    dob = df.loc[i][9]
    city = df.loc[i][10]
    dateTime = df.loc[i][11]
    country = df.loc[i][12]

    # Scores
    temp = []
    for j in range(13, 19):
        temp.append(df.loc[i][j])
    scorecard.append(temp)

    result = df.loc[i][19]
    if counter == noOfQs:
        records.append({'candNo': candNo, 'round': roundN, 'firstName': firstName, 'lastName': lastName, 'fullName': fullName,
                        'regNo': regNo, 'grade': grade, 'school': school, 'gender': gender, 'dob': dob, 'city': city,
                        'dateTime': dateTime, 'country': country, 'scorecard': scorecard, 'result': result})
        scorecard = []
        counter = 0
    counter += 1

# print(records)
for student in records:
    # Initializing FPDF object
    pdf = FPDF()
    pdf.add_page()

    # Headers
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 0, f'S. No. PQR / 2021 / 0{student["candNo"]}', 0, 0, 'L')

    pdf.set_font('Arial', 'B', 9)
    pdf.cell(0, 0, f'Regisration No.: {student["regNo"]}', 0, 0, 'R')
    pdf.ln(10)

    # Title of the Examination
    pdf.set_font('Times', 'B', 20)
    pdf.cell(
        0, 0, f'PQR Examination (Round {student["round"]}) - 2021', 0, 0, 'C')
    pdf.ln(8)

    pdf.set_font('Arial', '', 11)
    pdf.cell(
        0, 0, f'Class {student["grade"]} Grade Sheet cum Certificate of Performance', 0, 0, 'C')
    pdf.ln(8)

    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 0, 'MARKS STATEMENT', 0, 0, 'C')
    pdf.ln(10)
    
    # Image
    pdf.image(f'images/{student["fullName"]}.jpg', 140, 45, 40, 34)
    pdf.rect(138, 43, 44, 38)

    # Personal Information
    pdf.set_font('Arial', '', 10)
    pdf.cell(40, 0, 'Name of the Candidate: ', 0, 0, 'L')
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 0, f'{student["fullName"]}', 0, 0, 'L')
    pdf.ln(6)

    pdf.set_font('Arial', '', 10)
    pdf.cell(17, 0, 'Gender:', 0, 0, 'L')
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 0, f'{student["gender"]}', 0, 0, 'L')
    pdf.ln(6)

    pdf.set_font('Arial', '', 10)
    pdf.cell(25, 0, 'Date Of Birth:', 0, 0, 'L')
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 0, f'{student["dob"]}', 0, 0, 'L')
    pdf.ln(6)

    pdf.set_font('Arial', '', 10)
    pdf.cell(15, 0, 'School:', 0, 0, 'L')
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 0, f'{student["school"]} Sr. Sec. School', 0, 0, 'L')
    pdf.ln(6)

    pdf.set_font('Arial', '', 10)
    pdf.cell(17, 0, 'Address:', 0, 0, 'L')
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 0, f'{student["city"]}, {student["country"]}', 0, 0, 'L')
    pdf.ln(8)

    pdf.set_font('Arial', '', 10)
    pdf.cell(22, 0, 'has performed as follows: ', 0, 0, 'L')
    pdf.ln(5)
    pdf.set_font('Arial', 'I', 10)
    pdf.cell(
        22, 0, f'Academic Performance - Grade {student["grade"]}  Scholastic Areas', 0, 0, 'L')
    pdf.ln(5)

    # Scores - Table Format
    total_weightage = 0
    total_score = 0

    # Table Headers
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(31.5, 8, "Question No.", 1, 0, 'C')
    pdf.cell(31.5, 8, "You Marked", 1, 0, 'C')
    pdf.cell(31.5, 8, "Correct Answer", 1, 0, 'C')
    pdf.cell(31.5, 8, "Outcome", 1, 0, 'C')
    pdf.cell(31.5, 8, "Weightage", 1, 0, 'C')
    pdf.cell(31.5, 8, "Your Score", 1, 1, 'C')

    # Table Content
    pdf.set_font('Arial', '', 8)
    for det in student["scorecard"]:
        pdf.cell(31.5, 5, f"{det[0]}", 1, 0, 'C')
        pdf.cell(31.5, 5, f"{det[1]}", 1, 0, 'C')
        pdf.cell(31.5, 5, f"{det[2]}", 1, 0, 'C')
        pdf.cell(31.5, 5, f"{det[3]}", 1, 0, 'C')
        pdf.cell(31.5, 5, f"{det[4]}", 1, 0, 'C')
        pdf.cell(31.5, 5, f"{det[5]}", 1, 1, 'C')

        total_weightage += int(det[4])
        total_score += int(det[5])

    # Total
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(126, 6, "TOTAL", 1, 0, 'C')
    pdf.cell(31.5, 6, f"{total_weightage}", 1, 0, 'C')
    pdf.cell(31.5, 6, f"{total_score}", 1, 1, 'C')

    percentage = (total_score/total_weightage)*100
    pdf.cell(126, 6, "TOTAL PERCENTAGE", 1, 0, 'C')
    pdf.cell(63, 6, f"{percentage} %",  1, 1, 'C')

    # Other Details
    pdf.set_font('Arial', '', 8)
    pdf.ln(4)
    pdf.cell(45, 0, "*For any queries regarding performance and grades please contact your class teacher", 0, 1)
    pdf.ln(11)

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(80, 0, "Result: ", 0, 0, 'R')
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 0, f"{student['result']}")
    pdf.ln(21)

    # Footer Section
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(37, 0, "Date of Examination: ", 0, 0, 'L')

    pdf.set_font('Arial', '', 10)
    pdf.cell(85, 0, f"{student['dateTime']}", 0, 0, 'L')

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(50, 7, "Signature of Principal", 'T', 1, 'C')

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(27, 0, "Date of Result: ", 0, 0, 'L')

    pdf.set_font('Arial', '', 10)
    pdf.cell(100, 0, f"{date.today()}", 0, 0, 'L')

    pdf.output(f'results/{student["firstName"]}_result.pdf', 'F')