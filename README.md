# Registration-form-GUI
this is the project of registration form of GUI using python.

from openpyxl import load_workbook
from tkinter import *

# Load the workbook and sheet globally
wb = load_workbook("wb.xlsx")
sheet = wb.active

# Functions to set focus on next field
def focus0(event):
    course_field.focus_set()
def focus1(event):
    sem_field.focus_set()
def focus2(event):
    form_no_field.focus_set()
def focus3(event):
    contact_no_field.focus_set()
def focus4(event):
    email_id_field.focus_set()
def focus5(event):
    address_field.focus_set()

# Clear all input fields
def clear():
    name_field.delete(0, END)
    course_field.delete(0, END)
    sem_field.delete(0, END)
    form_no_field.delete(0, END)
    contact_no_field.delete(0, END)
    email_id_field.delete(0, END)
    address_field.delete(0, END)

# Insert data into Excel file
def insert():
    if (name_field.get() == "" and
        course_field.get() == "" and
        sem_field.get() == "" and
        form_no_field.get() == "" and
        contact_no_field.get() == "" and
        email_id_field.get() == "" and
        address_field.get() == ""):
        print("Empty Field")
    else:
        current_row = sheet.max_row
        sheet.cell(row=current_row + 1, column=1).value = name_field.get()
        sheet.cell(row=current_row + 1, column=2).value = course_field.get()
        sheet.cell(row=current_row + 1, column=3).value = sem_field.get()
        sheet.cell(row=current_row + 1, column=4).value = form_no_field.get()
        sheet.cell(row=current_row + 1, column=5).value = contact_no_field.get()
        sheet.cell(row=current_row + 1, column=6).value = email_id_field.get()
        sheet.cell(row=current_row + 1, column=7).value = address_field.get()
        wb.save("wb.xlsx")
        clear()
        name_field.focus_set()

if __name__ == "__main__":
    root = Tk()
    root.title("Registration Form")
    root.geometry("400x250")

    # Labels
    Label(root, text="Name", bg="light grey").grid(row=0, column=0)
    Label(root, text="Course", bg="light grey").grid(row=1, column=0)
    Label(root, text="Semester", bg="light grey").grid(row=2, column=0)
    Label(root, text="Form No.", bg="light grey").grid(row=3, column=0)
    Label(root, text="Contact Number", bg="light grey").grid(row=4, column=0)
    Label(root, text="Email-ID", bg="light grey").grid(row=5, column=0)
    Label(root, text="Address", bg="light grey").grid(row=6, column=0)

    # Entry fields
    name_field = Entry(root)
    course_field = Entry(root)
    sem_field = Entry(root)
    form_no_field = Entry(root)
    contact_no_field = Entry(root)
    email_id_field = Entry(root)
    address_field = Entry(root)

    # Bind Enter key to focus next field
    name_field.bind("<Return>", focus0)
    course_field.bind("<Return>", focus1)
    sem_field.bind("<Return>", focus2)
    form_no_field.bind("<Return>", focus3)
    contact_no_field.bind("<Return>", focus4)
    email_id_field.bind("<Return>", focus5)

    # Position Entry fields
    name_field.grid(row=0, column=1, ipadx=50)
    course_field.grid(row=1, column=1, ipadx=50)
    sem_field.grid(row=2, column=1, ipadx=50)
    form_no_field.grid(row=3, column=1, ipadx=50)
    contact_no_field.grid(row=4, column=1, ipadx=50)
    email_id_field.grid(row=5, column=1, ipadx=50)
    address_field.grid(row=6, column=1, ipadx=50)

    # Submit button
    submit = Button(root, text="Submit", fg="black", bg="blue", command=insert)
    submit.grid(row=7, column=1)

    name_field.focus_set()
    root.mainloop()

