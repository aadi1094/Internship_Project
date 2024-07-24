
import tkinter
from tkinter import ttk
from tkinter import messagebox
import tkinter.messagebox
import os
import openpyxl

# if __name__ == "__main__":
def Submit_Click():

    accepted=accept_vars.get()

    if accepted=="Accepted":
        firstname=first_name_entry.get()
        lastname=last_name_entry.get()

        if firstname and lastname:
            title=title_combobox.get()
            age=age_spinbox.get()
            nationality=nationality_combobox.get()

            registration_status=reg_status_var.get()
            numcourses=numcourses_spinbox.get()
            numsemesters=numsemesters_spinbox.get()


            print(f"First Name: {firstname}, Last Name: {lastname}, Title: {title}, Age: {age}, Nationality: {nationality}")
            print(f"Number of Courses: {numcourses}, Number of Semesters: {numsemesters}")
            print(f"Registration Status: {registration_status}") 
            print("-----------------------------------")

            #Saving the data to an excel file
            filepath="E:\ADITYA PYTHON\Tkinter\DataSet.xlsx"

            if not os.path.exists(filepath):
                wb=openpyxl.Workbook()
                sheet=wb.active
                heading=["First Name","Last Name","Title","Age","Nationality","Registration Status","Number of Courses","Number of Semesters"]
                sheet.append(heading)
                wb.save(filepath)
            wb=openpyxl.load_workbook(filepath)
            sheet=wb.active
            sheet.append([firstname,lastname,title,age,nationality,registration_status,numcourses,numsemesters])
            wb.save(filepath)


        else:
            tkinter.messagebox.showwarning(title="Error",message="Please enter your first and last name to proceed")
    else:
        tkinter.messagebox.showwarning(title="Error",message="Please accept the terms and conditions to proceed")
    



window=tkinter.Tk()
window.title("Data Entry Form")

frame=tkinter.Frame(window)
frame.pack()

#Saving the user information
user_info_frame=tkinter.LabelFrame(frame,text="User Information")
user_info_frame.grid(row=0,column=0,padx=20,pady=10)

first_name_label=tkinter.Label(user_info_frame,text="First Name")
first_name_label.grid(row=0,column=0)
last_name_label=tkinter.Label(user_info_frame,text="Last Name")
last_name_label.grid(row=0,column=1)

first_name_entry=tkinter.Entry(user_info_frame)
last_name_entry=tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1,column=0)
last_name_entry.grid(row=1,column=1)

title_label=tkinter.Label(user_info_frame,text="Title")
title_combobox=ttk.Combobox(user_info_frame,values=["","Mr.","Mrs.","Miss","Dr."])
title_label.grid(row=0,column=2)
title_combobox.grid(row=1,column=2)

age_label=tkinter.Label(user_info_frame,text="Age")
age_spinbox=tkinter.Spinbox(user_info_frame,from_=18,to=100)
age_label.grid(row=2,column=0)
age_spinbox.grid(row=3,column=0)

nationality_label=tkinter.Label(user_info_frame,text="Nationalality")
nationality_combobox=ttk.Combobox(user_info_frame,values=["","Indian","American","British","Chinese"])
nationality_label.grid(row=2,column=1)
nationality_combobox.grid(row=3,column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

#Saving course information
courses_frame=tkinter.LabelFrame(frame,text="Courses Information")
courses_frame.grid(row=1,column=0,sticky="news",padx=20,pady=10)

registered_label=tkinter.Label(courses_frame,text="Registered Courses")

reg_status_var=tkinter.StringVar(value="Not Registered")
registered_check=tkinter.Checkbutton(courses_frame,text="Currently Registered",variable=reg_status_var,onvalue="Registered",offvalue="Not Registered")
registered_label.grid(row=0,column=0)
registered_check.grid(row=1,column=0)

numcourses_label=tkinter.Label(courses_frame,text="Completed Courses")
numcourses_spinbox=tkinter.Spinbox(courses_frame,from_=0,to='infinity')
numcourses_label.grid(row=0,column=1)
numcourses_spinbox.grid(row=1,column=1)

numsemesters_label=tkinter.Label(courses_frame,text="Semesters Completed")
numsemesters_spinbox=tkinter.Spinbox(courses_frame,from_=0,to='infinity')
numsemesters_label.grid(row=0,column=2)
numsemesters_spinbox.grid(row=1,column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

#Accept terms and conditions

terms_frame=tkinter.LabelFrame(frame,text="Terms and Conditions")
terms_frame.grid(row=2,column=0,sticky="news",padx=20,pady=10)

accept_vars=tkinter.StringVar(value="Not Accepted")
terms_check=tkinter.Checkbutton(terms_frame,text="I accept the terms and conditions",variable=
                                accept_vars,onvalue="Accepted",offvalue="Not Accepted")
terms_check.grid(row=0,column=0)

#Submit button
submit_button=tkinter.Button(frame,text="Submit",command=Submit_Click)
submit_button.grid(row=3,column=0,pady=10,padx=20)





window.mainloop()