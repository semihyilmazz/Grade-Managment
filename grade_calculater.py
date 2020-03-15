from Tkinter import *
import tkMessageBox
import tkFileDialog
import xlrd
import anydbm
import pickle
from xlrd import open_workbook
import os.path


class GradeCalculator(Frame):
    def __init__(self,parent):
        self.parent = parent
        Frame.__init__(self,parent)
        self.initUI()


    def initUI(self):

        #Creating GUI and adding them Frames.


        self.Frame1 = Frame(self,bg ="light blue")
        self.Frame1.pack(fill=X)


        self.Frame2 = Frame(self)
        self.Frame2.pack(fill=X,pady=7)

        self.Framemidterm= Frame(self)
        self.Framemidterm.pack(fill=X,pady=7)

        self.FrameFinalGrading=Frame(self)
        self.FrameFinalGrading.pack(fill=X,pady=7)

        self.FrameAttendance=Frame(self)
        self.FrameAttendance.pack(fill=X,pady=7)

        self.FrameCalculateSave=Frame(self)
        self.FrameCalculateSave.pack(fill=X,pady=7,padx=175)

        self.FrameText=Frame(self)
        self.FrameText.pack(fill=X,pady=7)

        self.baslik = Label(self.Frame1,text="ENGR 102 Numerical Grade Calculator", bg="green")
        self.baslik.config(font=("Arial", 14))
        self.baslik.pack()



        #i get the percentange by Entrys
        #i gave entries int variable so when ever a input given it will be that input
        self.mp1var=IntVar()
        self.mp1 = Label(self.Frame2,text="         MP1 %  ")
        self.mp1.grid(row=0,column =0)
        self.mp1_result= Entry(self.Frame2,textvariable=self.mp1var,width=10)
        self.mp1_result.grid(row=0,column=1)

        self.mp2var=IntVar()
        self.mp2=Label(self.Frame2,text="  MP2 %  ")
        self.mp2.grid(row=0,column=2)
        self.mp2_result=Entry(self.Frame2,textvariable=self.mp2var,width=10)
        self.mp2_result.grid(row=0,column=3)

        self.mp3var=IntVar()
        self.mp3 = Label(self.Frame2,text="  MP3 %  ")
        self.mp3.grid(row=0, column=4)
        self.mp3_result=Entry(self.Frame2,textvariable=self.mp3var,width=10)
        self.mp3_result.grid(row=0,column=5)

        self.mp4var=IntVar()
        self.mp4 = Label(self.Frame2,text="  MP4 %  ")
        self.mp4.grid(row=0, column=6)
        self.mp4_result=Entry(self.Frame2,textvariable=self.mp4var,width=10)
        self.mp4_result.grid(row=0,column=7)

        self.mp5var=IntVar()
        self.mp5 = Label(self.Frame2,text="  MP5 %  ")
        self.mp5.grid(row=0, column=8)
        self.mp5_result=Entry(self.Frame2,textvariable=self.mp5var,width=10)
        self.mp5_result.grid(row=0,column=9)

        # i get the percentange by Entrys
        # i gave entries int variable so when ever a input given it will be that input

        self.midtermvar=IntVar()

        self.midterm= Label(self.Framemidterm,text="    Midterm %")
        self.midterm.grid(row=4,column=0)
        self.midterm_result=Entry(self.Framemidterm,textvariable=self.midtermvar,width=10)
        self.midterm_result.grid(row=4,column=1)

        self.finalvar=IntVar()
        self.final = Label(self.FrameFinalGrading,text="         Final %  ")
        self.final.grid(row=5,column=0)
        self.final_result=Entry(self.FrameFinalGrading,textvariable=self.finalvar,width=10)
        self.final_result.grid(row=5,column=1)


        self.attendance_result=Label(self.FrameAttendance,text="Attendance%")
        self.attendance_result.grid(row=6,column=0)

        self.attendancevar=IntVar()
        self.attendance_result_per=Entry(self.FrameAttendance,textvariable=self.attendancevar,width=10)
        self.attendance_result_per.grid(row=6,column=1)

        self.gradingfile= Label(self.FrameFinalGrading,text="             Grading File     ")
        self.gradingfile.grid(row=5,column=3)

        self.gradingbrowse= Button(self.FrameFinalGrading,text="Browse",command=self.gradingFile,bg="light pink")
        self.gradingbrowse.grid(row=5,column=5)

        self.attendance= Label(self.FrameAttendance,text="             Attendance      ")
        self.attendance.grid(row=6,column=3)

        self.attendancebrowse=Button(self.FrameAttendance,text="Browse",command=self.attendanceFile,bg="light pink")
        self.attendancebrowse.grid(row=6,column=5)

        self.calculate=Button(self.FrameCalculateSave,text="Calculate",command=self.calculate,bg="light pink")
        self.calculate.grid(row=7,column=2)

        self.save=Button(self.FrameCalculateSave,text="Save",command=self.save,bg="light pink")
        self.save.grid(row=7,column=5,padx=15)

        #so i used grid and row column to place everything correctly
        #created button to get excel files
        #created enteries to get percentanges


        self.all_text= Text(self.FrameText,width=75,height=15)       # text for showing names and grades
        self.all_text.grid()
        self.scrollbar = Scrollbar(self.FrameText, orient=VERTICAL)
        self.all_text.config(yscrollcommand=self.scrollbar.set)     # adding scroll bar to a text
        self.scrollbar.config(command=self.all_text.yview)
        self.scrollbar.grid(row =0,column =5,sticky = N+S)

        if os.path.exists("captions.db"):            # this if statement is for checking if there already database or not

            self.db = anydbm.open("captions.db","r")   # if there is database already then it'll be shown on text widget

            for key in self.db:

                self.all_text.insert(END,"   " + pickle.loads(key) + "     " + str(pickle.loads(self.db[key])) + "\n")

    def calculate(self):

        #its take if the given percentange is uqual to 100
        #by get() i sum all given percentages and see if its 100

        self.total=int(self.mp1_result.get())+int(self.mp2_result.get())+int(self.mp3_result.get())+int(self.mp4_result.get())+\
              int(self.mp5_result.get())+int(self.final_result.get())+int(self.attendance_result_per.get())+int(self.midterm_result.get())

        #if given percentange is uqual to 100 then it goos to else line:

        if self.total !=100:

            tkMessageBox.showinfo("Warning", "The assessment components do NOT sum up to 100.")

        else:
            self.all_text.delete("0.0",END)

            #so everytime user click calculate text widget will be clear and process will occur again
            #dicGrades is dict where i store names , mini projects scores, midterm and final and attendance grades
            #dicAT is dict where practice session grades stored

            self.dicGrades = {}
            self.dicAt={}

            for any in range(1, self.sheet.nrows):

                #i get name and surname by cell_value with a for loop

                self.name = self.sheet.cell_value(any, 2)
                self.surname = self.sheet.cell_value(any, 3)
                self.atgrade1 = 0

                #atgrade1 is lecture attendance grade

                for any2 in range(3, 17):

                    self.atgrade1 += int(self.sheet2.cell_value(any+1, any2))

                    #sum of lecture attendance grade for each person

                    self.dicGrades[self.name +" "+ self.surname] = [self.sheet.cell_value(any, 6),
                                                             self.sheet.cell_value(any, 7),
                                                             self.sheet.cell_value(any, 8),
                                                             self.sheet.cell_value(any, 9),
                                                             self.sheet.cell_value(any, 10),
                                                             self.sheet.cell_value(any, 11),
                                                             self.sheet.cell_value(any, 12), self.atgrade1,0,0]

                    # in dicGrades , keys = names and surname and values= their miniprojects midterm final and lecture attendace grades

                    #practice attendace

            for x in range(2,self.sheet3.nrows):

                self.atgrade2=0

                # atgrade 2  = practice session attendance grade

                self.name2 = self.sheet3.cell_value(x, 0)
                self.surname2 = self.sheet3.cell_value(x, 1)
                for y in range(3,17):

                    self.atgrade2+=int(self.sheet3.cell_value(x, y))
                self.dicAt[self.name2 +" "+ self.surname2] = [self.atgrade2]

                #self dicAt is dictionary where i put names as keys and practice session grades as values


            for every in self.dicGrades.keys():

                self.mp =self.dicGrades[every][0]* int(self.mp1_result.get())/100+self.dicGrades[every][1]* int(self.mp2_result.get())/100+self.dicGrades[every][2]* int(self.mp3_result.get())/100+self.dicGrades[every][3]* int(self.mp4_result.get())/100+self.dicGrades[every][4]* int(self.mp5_result.get())/100
                self.midfin=self.dicGrades[every][5]*int(self.midterm_result.get())/100 + self.dicGrades[every][6]*int(self.final_result.get())/100

                #calculating miniprojects (mp) and midterm final grades(midfin)

                if every in self.dicAt.keys():

                    #lecture attendace is in dicGrades and practice session attendace grade is in dicAt so
                    #i check 2 lists according to names and names, when it matches, i add practice sessin attendance grade to dicGrades
                    #so everything in dicGrades keys = names and values = [ mp1,mp2,mp3,mp4,mp5,midterm,final,lectureAt,practiceAt]

                    self.dicGrades[every][8] = self.dicAt[every][0]

                self.Atten = (self.dicGrades[every][7] + self.dicGrades[every][8]) * int(self.attendance_result_per.get())

                self.grade = self.mp + self.midfin + self.Atten / 28.0

                #as there is 28 attendance i calculate how much poin each person get


                self.dicGrades[every][9] = self.grade


                #adding last grades of all students in dicgrades[key][9] and that info will be used while operating database

                self.all_text.insert(END, "   "+every + "    " + str(self.grade) + "\n")



    def gradingFile(self):
        #getting files and opening here
        self.Gradefilename = tkFileDialog.askopenfilename(initialdir="/",
                                title="Select file",filetypes=(("jpeg files", "*.jpg"), ("all files", "*.*")))

        self.workbook = xlrd.open_workbook(self.Gradefilename)
        self.sheet = self.workbook.sheet_by_index(0)


    def attendanceFile(self):
        #getting files and opening here
        self.AttendancefileLocation = tkFileDialog.askopenfilename(initialdir="/",
                                title="Select file",filetypes=(("jpeg files", "*.jpg"), ("all files", "*.*")))

        self.workbook = xlrd.open_workbook(self.AttendancefileLocation)
        self.sheet2 = self.workbook.sheet_by_index(0)
        self.sheet3 =self.workbook.sheet_by_index(1)



    def save(self):
        self.db = anydbm.open('captions.db', 'c')

        for any in self.dicGrades.keys():

            self.db[pickle.dumps(any)] = pickle.dumps(self.dicGrades[any][9])


        #dicgrades[any][9] is students last grades so 271.line is a database and keys are names and values are last grades
        # i insert them to database
        ####################################################################

def main():
    root = Tk()
    root.configure(background='lightblue')
    app = GradeCalculator(root)
    app.grid()
    root.geometry("1000x650")

    root.mainloop()

main()