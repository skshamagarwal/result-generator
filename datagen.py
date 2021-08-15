from datetime import date
from fpdf import FPDF
from pandas import read_excel
from datetime import date
import matplotlib.pyplot as plt

df = read_excel('Dummy Data for final assignment.xlsx', sheet_name='Sheet1')

# Arranging Dummy Data in usable format
records = []
scorecard = []
counter = 1
noOfQs = 25
for i in range(1, len(df)):

    # Additional Details
    if counter == 1:
        averageScoreWorld = df.loc[i][24]
        medianScoreWorld = df.loc[i][25]
        modeScoreWorld = df.loc[i][26]
        firstNameAttempts = df.loc[i][27]
        averageAttempts = df.loc[i][28]
        firstNameAccuracy = round(df.loc[i][29]*100, 2)
        averageAccuracy = round(df.loc[i][30]*100, 2)
        additionalData = {'averageAccuracy': averageAccuracy, 'firstNameAccuracy': firstNameAccuracy, 'averageAttempts': averageAttempts, 'firstNameAttempts': firstNameAttempts, 'modeScoreWorld': modeScoreWorld,
                          'medianScoreWorld': medianScoreWorld, 'averageScoreWorld': averageScoreWorld}

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
    dob = list(str(df.loc[i][9]).split())[0]
    city = df.loc[i][10]
    dateTime = df.loc[i][11]
    country = df.loc[i][12]

    # Questions -
    # [0: Question No, 1: Marked, 2: Correct Answer, 3: Outcome, 4: Score if Correct, 5: Your Score,
    # 6: Percent world, 7: Percent Correct, 8: Percent Incorrect, 9: Average World]
    temp = []
    for j in range(13, 19):
        temp.append(df.loc[i][j])
    percentWorld = round(df.loc[i][20]*100, 2)
    percentCorrect = round(df.loc[i][21]*100, 2)
    percentIncorrect = round(df.loc[i][22]*100, 2)
    averageWorld = df.loc[i][23]
    temp.extend([percentWorld, percentCorrect, percentIncorrect, averageWorld])
    scorecard.append(temp)

    result = df.loc[i][19]
    if counter == noOfQs:
        records.append(dict({'candNo': candNo, 'round': roundN, 'firstName': firstName, 'lastName': lastName, 'fullName': fullName,
                             'regNo': regNo, 'grade': grade, 'school': school, 'gender': gender, 'dob': dob, 'city': city,
                             'dateTime': dateTime, 'country': country, 'scorecard': scorecard, 'result': result}, **additionalData))
        scorecard = []
        counter = 0
    counter += 1

# for i in records:
#     for j in i:
#         print(j, ":", i[j])
#     print('\n')
for student in records:
    # Initializing FPDF object 
    pdf = FPDF()
    pdf.add_page()

    # Background
    pdf.image(f'Utils 1-6/back.png', 0, 0, 220, 297)
    # Headers
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(
        54, 0, f'Round {student["round"]} - Enhanced Score Report: ', 0, 0, 'L')

    pdf.set_font('Arial', '', 9)
    pdf.cell(0, 0, f'{student["fullName"]}', 0, 2, 'L')
    pdf.ln(4)
    pdf.cell(0, 0, f'Regisration No.: {student["regNo"]}', 0, 0, 'L')
    pdf.ln(12)

    # Title of the Examination
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(
        155, 0, f'INTERNATIONAL MATHS OLYMPIAD CHALLENGE', 0, 0, 'C')
    pdf.ln(38)

    pdf.set_font('Arial', 'B', 12)
    pdf.cell(
        155, 0, f'Round {student["round"]} performance of {student["fullName"]}', 0, 0, 'C')
    pdf.ln(12)

    # Images
    # student image
    pdf.image(f'images/{student["fullName"]}.jpg', 155, 25, 44, 38)
    pdf.rect(155, 25, 44, 38)
    # logo image
    pdf.image(f'Utils 1-6/logo.png', 57, 30, 60, 30)

    # Personal Information
    pdf.cell(8, 0)

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(39, 6, "Grade", 1, 0, 'L')
    pdf.set_font('Arial', '', 10)
    pdf.cell(37, 6, f"{student['grade']}", 1, 0, 'L')

    pdf.cell(20, 0)

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(39, 6, "Registration No.", 1, 0, 'L')
    pdf.set_font('Arial', '', 10)
    pdf.cell(37, 6, f"{student['regNo']}", 1, 1, 'L')

    pdf.cell(8, 0)

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(39, 6, "School Name", 1, 0, 'L')
    pdf.set_font('Arial', '', 10)
    pdf.cell(37, 6, f"{student['school']}", 1, 0, 'L')

    pdf.cell(20, 0)

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(39, 6, "Gender", 1, 0, 'L')
    pdf.set_font('Arial', '', 10)
    pdf.cell(37, 6, f"{student['gender']}", 1, 1, 'L')

    pdf.cell(8, 0)

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(39, 6, "City of Residence", 1, 0, 'L')
    pdf.set_font('Arial', '', 10)
    pdf.cell(37, 6, f"{student['city']}", 1, 0, 'L')

    pdf.cell(20, 0)

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(39, 6, "Date of Birth", 1, 0, 'L')
    pdf.set_font('Arial', '', 10)
    pdf.cell(37, 6, f"{student['dob']}", 1, 1, 'L')

    pdf.cell(8, 0)

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(39, 6, "Country of Residence", 1, 0, 'L')
    pdf.set_font('Arial', '', 10)
    pdf.cell(37, 6, f"{student['country']}", 1, 0, 'L')

    pdf.cell(20, 0)

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(39, 6, "Date of Test", 1, 0, 'L')
    pdf.set_font('Arial', '', 10)
    pdf.cell(37, 6, f"{student['dateTime']}", 1, 1, 'L')
    pdf.ln(8)

    # Section 1
    pdf.set_font('Arial', 'BI', 13)
    pdf.cell(0, 0, "Section 1", 0, 0, 'C')
    pdf.ln(6)
    pdf.set_font('Arial', '', 10)
    pdf.cell(
        0, 0, f"This section describes {student['firstName']}'s performance v/s the Test in Grade {student['grade']}", 0, 0, 'C')
    pdf.ln(8)

    total_weightage = 0
    total_score = 0

    # Table Headers
    pdf.set_font('Arial', 'B', 10)
    pdf.set_text_color(256, 256, 256)
    pdf.cell(8, 0)
    pdf.cell(25, 8, "Question No.", 1, 0, 'C', 1)
    pdf.cell(25, 8, "Attempt Status", 1, 0, 'C', 1)
    pdf.cell(25, 8, f"{student['firstName']}'s Choice", 1, 0, 'C', 1)
    pdf.cell(25, 8, "Correct Answer", 1, 0, 'C', 1)
    pdf.cell(25, 8, "Outcome", 1, 0, 'C', 1)
    pdf.cell(25, 8, "Score if Correct", 1, 0, 'C', 1)
    pdf.cell(25, 8, f"{student['firstName']}'s' Score", 1, 1, 'C', 1)
    pdf.set_text_color(0, 0, 0)
    # Table Content
    pdf.set_font('Arial', '', 8)
    for det in student["scorecard"]:
        pdf.cell(8, 0)
        attemptStatus = ''
        if det[3] == "Correct" or det[3] == "Incorrect":
            attemptStatus = "Attempted"
        else:
            attemptStatus = "Unattempted"

        pdf.cell(25, 5, f"{det[0]}", 1, 0, 'C')
        pdf.cell(25, 5, f"{attemptStatus}", 1, 0, 'C')
        pdf.cell(25, 5, f"{det[1]}", 1, 0, 'C')
        pdf.cell(25, 5, f"{det[2]}", 1, 0, 'C')
        pdf.cell(25, 5, f"{det[3]}", 1, 0, 'C')
        pdf.cell(25, 5, f"{det[4]}", 1, 0, 'C')
        pdf.cell(25, 5, f"{det[5]}", 1, 1, 'C')
        total_weightage += int(det[4])
        total_score += int(det[5])

    # Total
    pdf.ln(6)
    pdf.set_font('Arial', 'BI', 10)
    pdf.cell(170, 0, "Total Score:", 0, 0, 'R')
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 0, f"{total_score}", 0, 1, 'L')
    pdf.ln(20)

    pdf.add_page()

    # Background
    pdf.image(f'Utils 1-6/back.png', 0, 0, 220, 297)
    # Section 2
    pdf.set_font('Arial', 'BI', 13)
    pdf.cell(0, 0, "Section 2", 0, 0, 'C')
    pdf.ln(6)
    pdf.set_font('Arial', '', 10)
    pdf.cell(
        0, 0, f"This section describes {student['firstName']}'s performance v/s Rest of the World in Grade {student['grade']}", 0, 0, 'C')
    pdf.ln(8)

    # Table Headers
    pdf.set_font('Arial', 'B', 9)
    pdf.set_text_color(256, 256, 256)
    pdf.cell(19, 12, "Question No.", 1, 0, 'L', 1)
    pdf.cell(19, 12, "Attempt Status", 1, 0, 'L', 1)
    pdf.cell(19, 12, f"{student['firstName']}'s Choice", 1, 0, 'L', 1)
    pdf.cell(19, 12, "Correct Answer", 1, 0, 'L', 1)
    pdf.cell(19, 12, "Outcome", 1, 0, 'L', 1)
    pdf.cell(19, 12, f"{student['firstName']}'s' Score", 1, 0, 'L', 1)
    pdf.cell(19, 12, f"% of students across the world who attempted this question", 1, 0, 'L', 1)
    pdf.cell(19, 12, f"% of students (from those who attempted this) who got it correct", 1, 0, 'L', 1)
    pdf.cell(19, 12, f"% of students (from those who attempted this) who got it incorrect", 1, 0, 'L', 1)
    pdf.cell(19, 12, f"World Average in this question", 1, 1, 'C', 1)
    pdf.set_text_color(0, 0, 0)

    # Table Content
    pdf.set_font('Arial', '', 8)
    for det in student["scorecard"]:
        attemptStatus = ''
        if det[3] == "Correct" or det[3] == "Incorrect":
            attemptStatus = "Attempted"
        else:
            attemptStatus = "Unattempted"
        pdf.cell(19, 5, f"{det[0]}", 1, 0, 'C')
        pdf.cell(19, 5, f"{attemptStatus}", 1, 0, 'C')
        pdf.cell(19, 5, f"{det[1]}", 1, 0, 'C')
        pdf.cell(19, 5, f"{det[2]}", 1, 0, 'C')
        pdf.cell(19, 5, f"{det[3]}", 1, 0, 'C')
        pdf.cell(19, 5, f"{det[5]}", 1, 0, 'C')
        pdf.cell(19, 5, f"{det[6]}", 1, 0, 'C')
        pdf.cell(19, 5, f"{det[7]}", 1, 0, 'C')
        pdf.cell(19, 5, f"{det[8]}", 1, 0, 'C')
        pdf.cell(19, 5, f"{det[9]}", 1, 1, 'C')
    pdf.ln()

    # table footer
    percentile = (total_score/total_weightage)*100
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(pdf.get_string_width(
        f"{student['firstName']}'s "), 0, f"{student['firstName']}'s ")
    pdf.set_font('Arial', '', 9)
    pdf.cell(pdf.get_string_width("overall percentile in the world is "),
            0, "overall percentile in the world is ")
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(pdf.get_string_width(f"{percentile}%ile"), 0, f"{percentile}%ile")
    pdf.set_font('Arial', '', 9)
    pdf.cell(pdf.get_string_width(f". This indicates that {student['firstName']} has more than "),
            0, f". This indicates that {student['firstName']} has more than ")
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(pdf.get_string_width(f"{percentile}%"), 0, f"{percentile}%")
    pdf.set_font('Arial', '', 9)
    pdf.cell(pdf.get_string_width(f" of students in the world and"),
            0, f" of students in the world and", 0, 1)
    pdf.ln(4)
    pdf.cell(pdf.get_string_width("lesser than "), 0, "lesser than ")
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(pdf.get_string_width(f"{100-percentile}%"), 0, f"{100-percentile}%")
    pdf.set_font('Arial', '', 9)
    pdf.cell(pdf.get_string_width(f" of students in the world."),
            0, f" of students in the world.")
    pdf.ln(10)

    # Overview
    pdf.set_font('Arial', '', 9)
    x = pdf.get_x()
    y = pdf.get_y()
    pdf.multi_cell(35, 4, "Average score of all students across the world", 1, 0, 'L')
    pdf.set_xy(x+35, y)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 12, f"{student['averageScoreWorld']}", 1, 0, 'L')

    pdf.cell(12, 0)

    pdf.set_font('Arial', '', 9)
    x = pdf.get_x()
    y = pdf.get_y()
    pdf.multi_cell(35, 4, f"{student['firstName']}'s attempts (Attempts x 100 / Total Questions)", 1, 0, 'L')
    pdf.set_xy(x+35, y)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 12, f"{student['firstNameAttempts']}%", 1, 0, 'L')

    pdf.cell(12, 0)

    pdf.set_font('Arial', '', 9)
    x = pdf.get_x()
    y = pdf.get_y()
    pdf.multi_cell(35, 4, f"{student['firstName']}'s Accuracy (Corrects x 100 /Attempts)", 1, 0, 'L')
    pdf.set_xy(x+35, y)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 12, f"{student['firstNameAccuracy']}%", 1, 1, 'L')

    pdf.set_font('Arial', '', 9)
    x = pdf.get_x()
    y = pdf.get_y()
    pdf.multi_cell(35, 4, "Median score of all students across the World", 1, 0, 'L')
    pdf.set_xy(x+35, y)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 12, f"{student['medianScoreWorld']}", 1, 0, 'L')

    pdf.cell(12, 0)

    pdf.set_font('Arial', '', 9)
    x = pdf.get_x()
    y = pdf.get_y()
    pdf.multi_cell(35, 4, "Average attempts of all students across the World", 1, 0, 'L')
    pdf.set_xy(x+35, y)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 12, f"{student['averageAttempts']}%", 1, 0, 'L')

    pdf.cell(12, 0)

    pdf.set_font('Arial', '', 9)
    x = pdf.get_x()
    y = pdf.get_y()
    pdf.multi_cell(35, 4, f"Average accuracy of all students across the world", 1, 0, 'L')
    pdf.set_xy(x+35, y)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 12, f"{student['averageAccuracy']}%", 1, 1, 'L')

    pdf.set_font('Arial', '', 9)
    x = pdf.get_x()
    y = pdf.get_y()
    pdf.multi_cell(35, 4, "Mode score of all students across World", 1, 0, 'L')
    pdf.set_xy(x+35, y)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(20, 8, f"{student['modeScoreWorld']}", 1, 0, 'L')

    # Plot Graphs
    # fig = plt.figure()
    # ax = fig.add_axes([0,0,1,1])
    # x_label = [f"{student['firstName']}, World"]
    # y_label = [student['averageAccuracy']]
    # ax.bar(x_label,y_label)
    # plt.show()

    pdf.ln(30)
    pdf.set_font('Arial', 'B', 13)
    pdf.cell(0, 0, f"Final Result: {student['result']}", 0, 0, 'C')

    pdf.output(f'results/{student["firstName"]}_{student["regNo"]}.pdf', 'F')
