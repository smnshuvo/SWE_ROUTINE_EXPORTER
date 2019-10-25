import xlrd
from tkinter import filedialog as fd
import tkinter
from tkinter import messagebox

# DEVELOPED BY SMNSHUVO
# Daffodil International University
# Using it without proper credential is considered against copyright law


# c1_code = "SE224" # TEST PURPOSE VARIABLE
# c2_code = "SE223"
# c3_code = "SE222"
# c4_code = "SE221"
# c5_code = "SE532"


# custom function
def routine_exporter(starting_cell, ending_cell, sheet, day):
    final_output = ''
    print("8.30 - 10.00")
    for x in range(starting_cell, ending_cell):
            temp = str(sheet.cell_value(x, 1))
            if section_checker(temp, section) and (
                    temp.__contains__(c1_code) or temp.__contains__(c2_code) or temp.__contains__(
                    c3_code) or temp.__contains__(c4_code) or temp.__contains__(c5_code)):
                roomNo = (sheet.cell_value(x, 0))  # ROOM NO
                courseCode = (sheet.cell_value(x, 1)),  # Course Code
                assignedTeacher = (sheet.cell_value(x, 2))  # Assigned Teacher
                output = "ROOM: {} COURSE: {} TEACHER -> {}"
                print(output.format(roomNo, courseCode, assignedTeacher))

                # Trying to make it an object
                # class 1 refers to first class
                final_output += "8.30 - 10.00 " + (output.format(roomNo, courseCode, assignedTeacher)) + "\n\n"

    print("10.00 - 11.30")
    for x in range (starting_cell, ending_cell):
            temp = str(sheet.cell_value(x, 4))
            if section_checker(temp, section) and (
                    temp.__contains__(c1_code) or temp.__contains__(c2_code) or temp.__contains__(
                    c3_code) or temp.__contains__(c4_code) or temp.__contains__(c5_code)):
                roomNo = (sheet.cell_value(x, 3))  # ROOM NO
                courseCode = (sheet.cell_value(x,4)), # Course Code
                assignedTeacher = (sheet.cell_value(x,5)) # Assigned Teacher
                output = "ROOM: {} COURSE: {} TEACHER -> {}"
                print(output.format(roomNo, courseCode, assignedTeacher))

                final_output += "10.00 - 11.30 " + (output.format(roomNo, courseCode, assignedTeacher)) + "\n\n"
    print("11.30 - 1.00")
    for x in range(starting_cell, ending_cell):
            temp = str(sheet.cell_value(x, 7))
            if section_checker(temp, section) and (
                    temp.__contains__(c1_code) or temp.__contains__(c2_code) or temp.__contains__(
                    c3_code) or temp.__contains__(c4_code) or temp.__contains__(c5_code)):
                roomNo = (sheet.cell_value(x, 6))  # ROOM NO
                courseCode = (sheet.cell_value(x, 7)),  # Course Code
                assignedTeacher = (sheet.cell_value(x, 8))  # Assigned Teacher
                output = "ROOM: {} COURSE: {} TEACHER -> {}"

                print(output.format(roomNo, courseCode, assignedTeacher))
                final_output += "11.30 - 1.00 " + (output.format(roomNo, courseCode, assignedTeacher)) + "\n\n"

    print("1.00 - 2.30")
    for x in range (starting_cell, ending_cell):
            temp = str(sheet.cell_value(x, 10))
            if section_checker(temp, section) and (
                    temp.__contains__(c1_code) or temp.__contains__(c2_code) or temp.__contains__(
                    c3_code) or temp.__contains__(c4_code) or temp.__contains__(c5_code)):
                roomNo = (sheet.cell_value(x, 9))  # ROOM NO
                courseCode = (sheet.cell_value(x, 10)),  # Course Code
                assignedTeacher = (sheet.cell_value(x, 11))  # Assigned Teacher
                output = "ROOM: {} COURSE: {} TEACHER -> {}"

                print(output.format(roomNo, courseCode, assignedTeacher))
                final_output += "1.00 - 2.30 " + (output.format(roomNo, courseCode, assignedTeacher)) + "\n\n"
    print("2.30 - 4.00")
    for x in range (starting_cell, ending_cell):
            temp = str(sheet.cell_value(x, 13))
            if section_checker(temp, section) and (
                    temp.__contains__(c1_code) or temp.__contains__(c2_code) or temp.__contains__(
                    c3_code) or temp.__contains__(c4_code) or temp.__contains__(c5_code)):
                roomNo = (sheet.cell_value(x, 12))  # ROOM NO
                courseCode = (sheet.cell_value(x, 13)),  # Course Code
                assignedTeacher = (sheet.cell_value(x, 14))  # Assigned Teacher
                output = "ROOM: {} COURSE: {} TEACHER -> {}"
                print(output.format(roomNo, courseCode, assignedTeacher))

                final_output += "2.30 - 4.00 " + (output.format(roomNo, courseCode, assignedTeacher)) + "\n\n"
    print("4.00 - 5.30")
    for x in range (starting_cell, ending_cell):
            temp = str(sheet.cell_value(x, 16))
            if section_checker(temp, section) and (
                    temp.__contains__(c1_code) or temp.__contains__(c2_code) or temp.__contains__(
                    c3_code) or temp.__contains__(c4_code) or temp.__contains__(c5_code)):
                roomNo = (sheet.cell_value(x, 15))  # ROOM NO
                courseCode = (sheet.cell_value(x, 16)),  # Course Code
                assignedTeacher = (sheet.cell_value(x, 17))  # Assigned Teacher
                output = "ROOM: {} COURSE: {} TEACHER -> {}"
                print(output.format(roomNo, courseCode, assignedTeacher))

                final_output += "4.00 - 5.30 " + (output.format(roomNo, courseCode, assignedTeacher)) + "\n\n"
    if final_output != '':
        # for x in range(len(classes)):
        # classes[x].get_info()
        # todays_routine = DailyRoutine(0, classes)
        messagebox.showinfo(day, final_output)

        # if there is no class on that day I don't wanna show that day


def student_routine_viewer():
    # Basically I have figured out that
    # the function works as a reference
    global c1_code
    global c2_code
    global c3_code
    global c4_code
    global c5_code
    global section
    # defined these as global variable
    c1_code = code1.get().upper()  # I know you can do mistakes :-)
    c2_code = code2.get().upper()
    c3_code = code3.get().upper()
    c4_code = code4.get().upper()
    c5_code = code5.get().upper()
    section = '' + sect.get().upper()
    # Give the location of the file
    loc = fd.askopenfilename()

    # Changed to GUI
    # To open Workbook
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    total_rows = sheet.nrows
    # For row 0 and column 0
    global saturday_limit, sunday_limit, monday_limit, wednesday_limit
    for x in range(total_rows):
        temp = str(sheet.cell_value(x, 0))
        # if temp.__contains__('SE2') or temp.__contains__('SWE2') :
        # print(sheet.cell_value(x, 0))

        if temp.__contains__("Sunday"):
            saturday_limit = x
            print("\033[4m      SATURDAY      \033[0m")
            routine_exporter(0, saturday_limit, sheet, "Saturday")

        if temp == "Monday":
            sunday_limit = x
            print("\033[4m      SUNDAY       \033[0m")
            routine_exporter(saturday_limit, sunday_limit, sheet, "Sunday")

        if temp == "Tuesday":
            monday_limit = x
            print("\033[4m      MONDAY       \033[0m")
            routine_exporter(sunday_limit, monday_limit, sheet, "Monday")

        if temp == "Wednesday":
            tuesday_limit = x
            print("\033[4m      TUESDAY       \033[0m")
            routine_exporter(monday_limit, tuesday_limit, sheet, "Tuesday")
        if temp.__contains__("Thursday"):  # I don't know why this doesn't work normally
            wednesday_limit = x
            print("\033[4m      WEDNESDAY       \033[0m")
            routine_exporter(tuesday_limit, wednesday_limit, sheet, "Wednesday")

            print("\033[4m      THURSDAY       \033[0m")
            routine_exporter(wednesday_limit, total_rows, sheet, "Thursday")


def section_checker(input, sec):
    if input.__contains__(' ' + sec):  # SE232 A
        return True
    elif input.__contains__(sec + '_'):  # SE232A_LAB1
        return True
    elif input.endswith(sec):
        return True
    return False


# Main method
m = tkinter.Tk()
m.title('SWE ROUTINE EXPORTER')
m.geometry("500x500")

m.resizable(0, 0)
code1 = tkinter.StringVar()
code2 = tkinter.StringVar()
code3 = tkinter.StringVar()
code4 = tkinter.StringVar()
code5 = tkinter.StringVar()
sect = tkinter.StringVar()
tkinter.Label(m, text="Course Code   #1").grid(row=0)
tkinter.Label(m, text="Course Code   #2").grid(row=1)
tkinter.Label(m, text="Course Code   #3").grid(row=2)
tkinter.Label(m, text="Course Code   #4").grid(row=3)
tkinter.Label(m, text="Course Code   #5").grid(row=4)
tkinter.Label(m, text="Section").grid(row=5)
tkinter.Label(m, text="Don't leave any of the field blank.").grid(row=6)
tkinter.Label(m, text="Use 'null' if the course code is empty").grid(row=7)
tkinter.Label(m, text="Developed by").grid(row=8)
tkinter.Label(m, text="SMN Shuvo").grid(row=9)

c1 = tkinter.Entry(m, textvariable=code1)
c2 = tkinter.Entry(m, textvariable=code2)
c3 = tkinter.Entry(m, textvariable=code3)
c4 = tkinter.Entry(m, textvariable=code4)
c5 = tkinter.Entry(m, textvariable=code5)
sec = tkinter.Entry(m, textvariable=sect)
c1.grid(row=0, column=1)
c2.grid(row=1, column=1)
c3.grid(row=2, column=1)
c4.grid(row=3, column=1)
c5.grid(row=4, column=1)
sec.grid(row=5, column=1)


button = tkinter.Button(m, text='Set and Go', width=15, command=student_routine_viewer)
teacher_button = tkinter.Button(m, text='View as Teacher', width=15, command=m.destroy)
#  button.pack()
button.grid(row=6, column=1)
tkinter.Label(m, text="If you are a teacher").grid(row=11)
tkinter.Label(m, text="You can also find your classes").grid(row=12)
teacher_button.grid(row=13)

m.mainloop()






