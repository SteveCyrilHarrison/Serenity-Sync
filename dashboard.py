from tkinter import *
import os
from tkinter import messagebox,ttk
from PIL import Image,ImageTk
from datetime import datetime
from tkcalendar import DateEntry
import random,membership_backend,giving_backend,attendace_backend
from tkinter import filedialog
import sqlite3
from reportlab.lib.pagesizes import A3
from reportlab.platypus import SimpleDocTemplate, Table,Paragraph,Spacer
from reportlab.lib.styles import getSampleStyleSheet
import subprocess
import io, cv2
import pandas as pd,csv
import matplotlib.pyplot as plt
import requests
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.dates import date2num
import webbrowser


class CManagement:
    def __init__(self,window):
        self.window = window
        self.window.geometry("1360x760+0+0")
        self.window.title("SerentySync")
        self.window.resizable(0,0)
        self.window.after(0,self.update_count)
        self.user = StringVar()
        self.user.set("USER-ADMIN")
        self.allmem =StringVar()
        self.men = StringVar()
        self.female = StringVar()
        self.bap = StringVar()
        

        # calling Functions
        
        self.DashBoard()
        self.TimeDisplay()

        # Variables for Attendance 
        
        self.attend_date = StringVar()
        self.attend_program = StringVar()
        self.attend_male = IntVar()
        self.attend_female = IntVar()
        self.attend_total = IntVar()
        self.attend_youth = IntVar()
        self.attend_child = IntVar()
        self.search_string_attend = StringVar()

        self.window.after(0,self.update_count)
        # self.master.after(0, self.update_counts)
        self.colors = ['#8e44ad', 'orange','#008080',"#5d6d7e",'cyan','purple'] 
        self.colors_bap = ['violet','green']   
        # self.window.after(0,self.attendanceCalculator)

    def graph(self):
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        cur.execute("SELECT department, COUNT(*) FROM members GROUP BY department")
        data = cur.fetchall()
        conn.close()
        if data:
            bap, population = zip(*data)
            self.ax1.clear()
            explode = [0.1 if department == 'Children' else 0 for department in bap]

            # Create the pie chart with labels and colors for each level
            self.ax1.pie(population,explode = explode ,labels=bap, colors=self.colors,
                        autopct='%1.1f%%', shadow=True, startangle=90)

            self.ax1.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.
            

            self.ax1.set_title('DEPARTMENTS ', fontsize=14, fontweight='bold')
            self.ax1.legend(title='department', loc='upper left', labels=bap, bbox_to_anchor=(0, 0.3), fontsize='medium')
        else:
            # If there is no data, clear the plot
            self.ax1.clear()
            self.ax1.set_title('No Data', fontsize=14, fontweight='bold')
            self.ax1.legend().remove()

    def lineGraph(self):
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        cur.execute("SELECT baptised, COUNT(*) FROM members GROUP BY baptised")
        data = cur.fetchall()
        conn.close()
        if data:
            bap, pop = zip(*data)
            self.ax2.clear()
            explode = [0.08 if bapt == 'No' else 0 for bapt in bap]

            # Create the pie chart with labels and colors for each level
            self.ax2.pie(pop,explode = explode ,labels=bap, colors=self.colors_bap,
                        autopct='%1.1f%%', shadow=True, startangle=230)

            self.ax2.axis('equal')  # Equal aspect ratio ensures the pie chart is circular.
            

            self.ax2.set_title('BAPTISM', fontsize=14, fontweight='bold')
            self.ax2.legend(title='Baptism', loc='upper left', labels=bap, bbox_to_anchor=(0.01, 0.3), fontsize='medium')
        else:
            # If there is no data, clear the plot
            self.ax2.clear()
            self.ax2.set_title('No Data', fontsize=14, fontweight='bold')
            self.ax2.legend().remove()

    def update_graph(self):
        self.ax1.clear()
        self.graph()
        self.count_male()
        self.count_female()
        self.count_Baptised()
        self.count_all()
        
        # Update the canvas and schedule the next update
        self.canvas1.draw()
        self.card_piechart_frame.after(2000, self.update_graph)
        # self.card_graph_bar.after(2000, self.update_graph)

    def update_graph_2(self):
        self.ax2.clear()
        self.lineGraph()
        self.count_male()
        self.count_female()
        self.count_Baptised()
        self.count_all()

        self.canvas.draw()
        self.card_graph_bar.after(2000, self.update_graph_2)


    def update_count(self):
        self.count_male()
        self.count_female()
        self.count_Baptised()
        self.count_all()
        self.card_piechart_frame.after(2000, self.update_graph)
        self.card_graph_bar.after(2000, self.update_graph_2)
        
    # excel Sheet
    def sheet(self):
        cols = ["S/N", "Membership ID",  "Full Name","Date Of Birth","Gender", "Confirm Baptism","Residence", "Department","Phone","Nationality","Email"]  # Your column headings here
        path = 'read.csv'
        excel_name = 'membership{}.xlsx'.format(random.randint(1,1000))
        lst = []
        with open(path, "w", newline='') as myfile:
            csvwriter = csv.writer(myfile,delimiter=',')
            for row_id in self.treeview.get_children():
                row = self.treeview.item(row_id, 'values')
                lst.append(row)
            lst = list(map(list, lst))
            lst.insert(0, cols)
            for row in lst:
                csvwriter.writerow(row)

        writer = pd.ExcelWriter(excel_name)
        df = pd.read_csv(path)
        df.to_excel(writer, 'sheetname{}'.format(random.randint(1,10)))
        writer._save()
        messagebox.showinfo("export message","Data Exported Successfully",parent = self.allwin)

   

    # Giving
    def sheet_Give(self):
        self.current_date = datetime.now()
        self.formatted_date = self.current_date.strftime("%A, %d %B, %Y")
        cols = ["S/N", "Type of Offering",  "Amount","Date Of Payment","Service"]
        path = 'read.csv'
        excel_name = 'GivingSheet_{}.xlsx'.format(self.formatted_date)
        lst = []
        with open(path, "w", newline='') as myfile:
            csvwriter = csv.writer(myfile,delimiter=',')
            for row_id in self.treeview.get_children():
                row = self.treeview.item(row_id, 'values')
                lst.append(row)
            lst = list(map(list, lst))
            lst.insert(0, cols)
            for row in lst:
                csvwriter.writerow(row)

        writer = pd.ExcelWriter(excel_name)
        df = pd.read_csv(path)
        df.to_excel(writer, 'sheetname_{}'.format(self.formatted_date))
        writer._save()
        messagebox.showinfo("export message","Data Exported Successfully",parent = self.givewin)

    # attend
    def sheet_attend(self):
        self.current_date = datetime.now()
        self.formatted_date = self.current_date.strftime("%A, %d %B, %Y")
        cols = ["S/N", "Date of Attendance",  "Program","Number Of Males","Number Of Females","Total Attendance"]
        path = 'read.csv'
        excel_name = 'AttendanceSheet_{}.xlsx'.format(self.formatted_date)
        lst = []
        with open(path, "w", newline='') as myfile:
            csvwriter = csv.writer(myfile,delimiter=',')
            for row_id in self.treeview.get_children():
                row = self.treeview.item(row_id, 'values')
                lst.append(row)
            lst = list(map(list, lst))
            lst.insert(0, cols)
            for row in lst:
                csvwriter.writerow(row)

        writer = pd.ExcelWriter(excel_name)
        df = pd.read_csv(path)
        df.to_excel(writer, 'sheetname{}'.format(self.formatted_date))
        writer._save()
        messagebox.showinfo("export message","Data Exported Successfully",parent =self.attendwin)

    # delete from DataBase
    
    def delete(self):
        if self.treeview.selection():
            result = messagebox.askquestion('Python - Delete Data Row In SQLite',
                                            'Are you sure you want to delete this record?', icon="warning",parent = self.allwin)
            if result == 'yes':
                curItem = self.treeview.focus()
                contents = (self.treeview.item(curItem))
                selecteditem = contents['values']
                self.treeview.delete(curItem)
                membership_backend.delete_member(selecteditem[0])

                self.DisplayData_members()

            else:
                self.DisplayData_members()

    def delete_tithe(self):
        if self.treeview.selection():
            result = messagebox.askquestion('Python - Delete Data Row In SQLite',
                                            'Are you sure you want to delete this record?', icon="warning",parent = self.givewin)
            if result == 'yes':
                curItem = self.treeview.focus()
                contents = (self.treeview.item(curItem))
                selecteditem = contents['values']
                self.treeview.delete(curItem)
                giving_backend.delete_Tithe(selecteditem[0])

                self.DisplayData_Tithe()

            else:
                self.DisplayData_Tithe()

    def delete_attend(self):
        if self.treeview.selection():
            result = messagebox.askquestion('Python - Delete Data Row In SQLite',
                                            'Are you sure you want to delete this record?', icon="warning",parent = self.attendwin)
            if result == 'yes':
                curItem = self.treeview.focus()
                contents = (self.treeview.item(curItem))
                selecteditem = contents['values']
                self.treeview.delete(curItem)
                attendace_backend.delete_attend(selecteditem[0])

                self.DisplayData_attend()

            else:
                self.DisplayData_attend()

    def delete_Give(self):
        if self.treeview.selection():
            result = messagebox.askquestion('Python - Delete Data Row In SQLite',
                                            'Are you sure you want to delete this record?', icon="warning",parent = self.givewin)
            if result == 'yes':
                curItem = self.treeview.focus()
                contents = (self.treeview.item(curItem))
                selecteditem = contents['values']
                self.treeview.delete(curItem)
                giving_backend.delete_Give(selecteditem[0])

                self.DisplayData_Give()

            else:
                self.DisplayData_Give()

    # Generate Pdf from Database
    def generate_pdf(self):
        doc = SimpleDocTemplate("treeview.pdf", pagesize=A3)
        data = []
        columns = []
        for col in self.treeview["columns"]:
            column_heading = self.treeview.heading(col)["text"]
            columns.append(column_heading)
        data.append(columns)
        for item in self.treeview.get_children():
            values = self.treeview.item(item)["values"]
            data.append(values)
        table_data = Table(data)
        styles = getSampleStyleSheet()
        heading_style = styles["Heading1"]
        heading_style.alignment = 1 
        heading = Paragraph("<b>DATASHEET</b>", heading_style)

        elements = [heading, Spacer(1, 20), table_data]
        doc.build(elements)
        subprocess.run(["start", "treeview.pdf"], shell=True)


    def generate_pdf_give(self):
        doc = SimpleDocTemplate("treeview.pdf")
        data = []
        columns = []
        for col in self.treeview["columns"]:
            column_heading = self.treeview.heading(col)["text"]
            columns.append(column_heading)
        data.append(columns)
        for item in self.treeview.get_children():
            values = self.treeview.item(item)["values"]
            data.append(values)
        table_data = Table(data)
        styles = getSampleStyleSheet()
        heading_style = styles["Heading1"]
        heading_style.alignment = 1 
        heading = Paragraph("<b>DATASHEET</b>", heading_style)

        elements = [heading, Spacer(1, 20), table_data]
        doc.build(elements)
        subprocess.run(["start", "treeview.pdf"], shell=True)


    def search_data_Give(self):
        search_string_data = self.search_string_give.get() 
        if search_string_data:
            self.treeview.delete(*self.treeview.get_children())
            conn = sqlite3.connect('GiveDB.db')
            query = f"SELECT * FROM giving WHERE offtype LIKE ? OR amount LIKE ? OR date LIKE ? OR service LIKE ?"
            cursor = conn.execute(query, ('%' + search_string_data + '%', '%' + search_string_data + '%', '%' + search_string_data + '%', '%' + search_string_data + '%'))
            fetch = cursor.fetchall()
            count = 0
            for data in fetch:
                if count % 2 == 0:
                    self.treeview.insert('', 'end', values=data, tags=('evenrow'))
                else:
                    self.treeview.insert('', 'end', values=data, tags=('oddrow'))
                count += 1
            cursor.close()
            conn.close()

    # Search Data from Database
    def search_data(self):
        search_string_data = self.search_string.get() 
        if search_string_data:
            self.treeview.delete(*self.treeview.get_children())
            conn = sqlite3.connect('database.db')
            query = f"SELECT * FROM members WHERE memID LIKE ? OR name LIKE ? OR gender LIKE ? OR department LIKE ?"
            cursor = conn.execute(query, ('%' + search_string_data + '%', '%' + search_string_data + '%', '%' + search_string_data + '%', '%' + search_string_data + '%'))
            fetch = cursor.fetchall()
            count = 0
            for data in fetch:
                if count % 2 == 0:
                    self.treeview.insert('', 'end', values=data, tags=('evenrow'))
                else:
                    self.treeview.insert('', 'end', values=data, tags=('oddrow'))
                count += 1
            cursor.close()
            conn.close()

    # Church Dashboard
    def DashBoard(self):
        self.win_frame = Frame(self.window,width=1360,height = 760,bg = "#f5f5f5")
        self.win_frame.pack()

        self.side_frame = Frame(self.win_frame,width = 270,height = 700,bg= '#008080')
        self.side_frame.place(relx= 0,rely =.08)
        self.contentFrame = Frame(self.win_frame,width =1090,height =680)
        self.contentFrame.place(relx = .2,rely =.104 )
        
        self.navigation_bar = Frame(self.win_frame,width = 1360,height = 80,bg ='white')
        self.navigation_bar.place(relx=0,rely =0)
        
        self.mainlogo = Label(self.navigation_bar,text = "SerenitySync",font = ('Montserrat',19,'bold'),bg = "white",fg = "#008080")
        self.mainlogo.place(relx = .09,rely = .4)
        self.excel101 = Button(self.navigation_bar,bg = 'white',bd = 0)
        self.logo101 = Image.open('img/logo12.png')
        self.resized101= self.logo101.resize((100,60))
        self.real_image_101 =ImageTk.PhotoImage(self.resized101)
        self.excel101.config(image =self.real_image_101)
        self.excel101.place(relx = 0.01,rely= 0.18)


        self.pButton = Button(self.navigation_bar,bg = 'white',bd = 0,command = self.closeMain)
        self.logo2 = Image.open('img/logout.png')
        self.resized2= self.logo2.resize((25,30))
        self.real_image2 =ImageTk.PhotoImage(self.resized2)
        self.pButton.config(image =self.real_image2)
        self.pButton.place(relx = 0.94,rely= 0.3)

        self.pButton = Button(self.navigation_bar,bg = 'white',bd = 0)
        self.logou = Image.open('img/user.png')
        self.resizedu= self.logou.resize((30,30))
        self.real_imageu =ImageTk.PhotoImage(self.resizedu)
        self.pButton.config(image =self.real_imageu)
        self.pButton.place(relx = 0.87,rely= 0.25)

        self.userEntry = Entry(self.navigation_bar,font = ('Montserrat',10,'bold'),width = 11,state = 'disabled',disabledbackground="white",bd = 0,disabledforeground="#008080",textvariable=self.user)
        self.userEntry.place(relx=.857,rely=.7)       

        self.card1 = Frame(self.contentFrame,width =250,height =150,bg = '#5d6d7e')
        self.card1.place(relx = .02,rely = .06)

        self.cardButton = Button(self.card1,bg = '#5d6d7e',bd = 0)
        self.card_1 = Image.open('img/all.png')
        self.resized_1= self.card_1.resize((80,80))
        self.real_image_1 =ImageTk.PhotoImage(self.resized_1)
        self.cardButton.config(image =self.real_image_1)
        self.cardButton.place(relx = 0.1,rely= 0.3)

        self.alllbl = Label(self.card1,text = "Total Members",font = ('Montserrat',11,'bold'),bg ='#5d6d7e',fg = "white")
        self.alllbl.place(relx=.48,rely = .7)

        self.allEntry = Entry(self.card1,font = ('Montserrat',28,'bold'),width=3,textvariable=self.allmem,state='disabled',disabledbackground="#5d6d7e",disabledforeground="white",bd = 0)
        self.allEntry.place(relx=.53,rely =.28)

        self.card2 = Frame(self.contentFrame,width =250,height =150,bg = '#3498db')
        self.card2.place(relx = .26,rely = .06)

        self.menlbl = Label(self.card2,text = "Total Males",font = ('Montserrat',11,'bold'),bg ='#3498db',fg = "white")
        self.menlbl.place(relx=.48,rely = .7)

        self.menEntry = Entry(self.card2,font = ('Montserrat',28,'bold'),width=3,textvariable=self.men,state='disabled',disabledbackground="#3498db",disabledforeground="white",bd = 0)
        self.menEntry.place(relx=.53,rely =.28)

        self.cardButton = Button(self.card2,bg = '#3498db',bd = 0)
        self.card_2 = Image.open('img/men.png')
        self.resized_2= self.card_2.resize((80,80))
        self.real_image_2 =ImageTk.PhotoImage(self.resized_2)
        self.cardButton.config(image =self.real_image_2)
        self.cardButton.place(relx = 0.1,rely= 0.3)

        self.card3 = Frame(self.contentFrame,width =250,height =150,bg = '#8e44ad')
        self.card3.place(relx = .5,rely = .06)

        self.femalelbl = Label(self.card3,text = "Total Females",font = ('Montserrat',11,'bold'),bg ='#8e44ad',fg = "white")
        self.femalelbl.place(relx=.48,rely = .7)

        # all females entry
        self.femaleEntry = Entry(self.card3,font = ('Montserrat',28,'bold'),width=3,textvariable=self.female,state='disabled',disabledbackground="#8e44ad",disabledforeground="white",bd = 0)
        self.femaleEntry.place(relx=.53,rely =.28)

        self.cardButton = Button(self.card3,bg = '#8e44ad',bd = 0)
        self.card_3 = Image.open('img/female.png')
        self.resized_3= self.card_3.resize((80,80))
        self.real_image_3 =ImageTk.PhotoImage(self.resized_3)
        self.cardButton.config(image =self.real_image_3)
        self.cardButton.place(relx = 0.1,rely= 0.3)

        self.card4 = Frame(self.contentFrame,width =250,height =150,bg = '#004d40')
        self.card4.place(relx = .74,rely = .06)

        self.baplbl = Label(self.card4,text = "Total baptised",font = ('Montserrat',11,'bold'),bg ='#004d40',fg = "white")
        self.baplbl.place(relx=.48,rely = .7)

        # Baptismal entry

        self.bapEntry = Entry(self.card4,font = ('Montserrat',28,'bold'),width=3,textvariable=self.bap,state='disabled',disabledbackground="#004d40",disabledforeground="white",bd = 0)
        self.bapEntry.place(relx=.53,rely =.28)

        self.cardButton = Button(self.card4,bg = '#004d40',bd = 0)
        self.card_4 = Image.open('img/bap.png')
        self.resized_4= self.card_4.resize((75,80))
        self.real_image_4 =ImageTk.PhotoImage(self.resized_4)
        self.cardButton.config(image =self.real_image_4)
        self.cardButton.place(relx = 0.1,rely= 0.3)

        # graph cards for graphing
        self.card_graph_bar = Frame(self.contentFrame,width =620,height =400,bg = '#f5f5f5')
        self.card_graph_bar.place(relx = .02,rely = .3)
        self.card_graph_bar.after(5000, self.update_graph)
        self.figure = plt.Figure(figsize=(6,4), dpi=100, facecolor="white")
        self.ax1 = self.figure.add_subplot(111)
        self.canvas1 = FigureCanvasTkAgg(self.figure, master=self.card_graph_bar)
        self.canvas1.get_tk_widget().place(relx = 0,rely = 0)

        self.card_piechart_frame = Frame(self.contentFrame,width =380,height =300,bg = '#008080')
        self.card_piechart_frame.place(relx = .62,rely = .44)
        self.card_piechart_frame.after(5000, self.update_graph_2)
        self.figure = plt.Figure(figsize=(4,3), dpi=100, facecolor="white")
        self.ax2 = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.card_piechart_frame)
        self.canvas.get_tk_widget().place(relx = 0,rely = 0)

        # instant messaging and voice assistance
        self.card_cal = Frame(self.contentFrame,width =180,height =80,bg = '#3498db')
        self.card_cal.place(relx = .62,rely = .3)

        self.cardButton = Button(self.card_cal,bg = '#3498db',bd = 0,command = self.calendar)
        self.card_c = Image.open('img/cal.png')
        self.resized_c= self.card_c.resize((80,60))
        self.real_image_c =ImageTk.PhotoImage(self.resized_c)
        self.cardButton.config(image =self.real_image_c)
        self.cardButton.place(relx = 0.2,rely= 0.1)

        self.card_voice_a = Frame(self.contentFrame,width =180,height =80,bg = 'green')
        self.card_voice_a.place(relx = .8,rely = .3)

        self.cardButton = Button(self.card_voice_a,bg = 'green',bd = 0,command = self.AiFun)
        self.card_v = Image.open('img/assis.png')
        self.resized_v= self.card_v.resize((80,60))
        self.real_image_v =ImageTk.PhotoImage(self.resized_v)
        self.cardButton.config(image =self.real_image_v)
        self.cardButton.place(relx = 0.23,rely= 0.1)

        self.image_lbl = Label(self.side_frame,bg = '#008080')
        self.logo = Image.open('img/new.png')
        self.resized = self.logo.resize((40,40))
        self.real_image =ImageTk.PhotoImage(self.resized)
        self.image_lbl.config(image =self.real_image)
        self.image_lbl.place(relx = 0.04,rely= 0.14)

        self.lbl1 = Button(self.side_frame,text = "New Membership",font = ('Montserrat',15),bg = '#008080',fg = 'white',bd =0,command = self.addMember)
        self.lbl1.place(relx =.22,rely = .14)

        self.image_lbl = Label(self.side_frame,bg = '#008080')
        self.logo1 = Image.open('img/db.png')
        self.resized1= self.logo1.resize((40,40))
        self.real_image1 =ImageTk.PhotoImage(self.resized1)
        self.image_lbl.config(image =self.real_image1)
        self.image_lbl.place(relx = 0.04,rely= 0.24)
        

        self.lbl2 = Button(self.side_frame,text = "All Members Data",font = ('Montserrat',14),bg = '#008080',fg = 'white',bd =0,command=self.allMembersData)
        self.lbl2.place(relx =.22,rely = .24)

        self.image_lbl = Label(self.side_frame,bg = '#008080')
        self.logo3 = Image.open('img/give.png')
        self.resized3= self.logo3.resize((40,40))
        self.real_image3 =ImageTk.PhotoImage(self.resized3)
        self.image_lbl.config(image =self.real_image3)
        self.image_lbl.place(relx = 0.04,rely= 0.34)

        self.lbl3 = Button(self.side_frame,text = "Giving",font = ('Montserrat',14),bg = '#008080',fg = 'white',bd =0,command = self.Giving)
        self.lbl3.place(relx =.22,rely = .34)

        self.image_lbl = Label(self.side_frame,bg = '#008080')
        self.logo4 = Image.open('img/msg.png')
        self.resized4= self.logo4.resize((40,40))
        self.real_image4 =ImageTk.PhotoImage(self.resized4)
        self.image_lbl.config(image =self.real_image4)
        self.image_lbl.place(relx = 0.04,rely= 0.44)

        self.lbl4 = Button(self.side_frame,text = "Instant Messaging",font = ('Montserrat',14),bg = '#008080',fg = 'white',bd =0,command = self.instantMessage)
        self.lbl4.place(relx =.22,rely = .44)

        self.image_lbl = Label(self.side_frame,bg = '#008080')
        self.logo5 = Image.open('img/attend.png')
        self.resized5= self.logo5.resize((40,40))
        self.real_image5 =ImageTk.PhotoImage(self.resized5)
        self.image_lbl.config(image =self.real_image5)
        self.image_lbl.place(relx = 0.04,rely= 0.54)

        self.lbl5 = Button(self.side_frame,text = "Attendance ",font = ('Montserrat',14),bg = '#008080',fg = 'white',bd =0,command = self.attendance)
        self.lbl5.place(relx =.22,rely = .54)

        self.image_lbl = Label(self.side_frame,bg = '#008080')
        self.logo6 = Image.open('img/set.png')
        self.resized6= self.logo6.resize((40,40))
        self.real_image6 =ImageTk.PhotoImage(self.resized6)
        self.image_lbl.config(image =self.real_image6)
        self.image_lbl.place(relx = 0.04,rely= 0.64)

        self.lbl3 = Button(self.side_frame,text = "Settings",font = ('Montserrat',14),bg = '#008080',fg = 'white',bd =0,command = "self.setting")
        self.lbl3.place(relx =.22,rely = .64)

        self.image_lbl = Label(self.side_frame,bg = '#008080')
        self.logo7 = Image.open('img/dev.png')
        self.resized7= self.logo7.resize((40,40))
        self.real_image7 =ImageTk.PhotoImage(self.resized7)
        self.image_lbl.config(image =self.real_image7)
        self.image_lbl.place(relx = 0.04,rely= 0.74)

        self.lbl3 = Button(self.side_frame,text = "Developers",font = ('Montserrat',14),bg = '#008080',fg = 'white',bd =0,command = self.Developers)
        self.lbl3.place(relx =.22,rely = .74)

    def addMember(self):
        self.addwin = Toplevel()
        self.addwin.geometry('1095x680+270+111')
        self.addwin.config(bg ='white')
        self.addwin.overrideredirect(True)

        self.memID = StringVar()
        self.name = StringVar()
        self.dob = StringVar()
        self.gender = StringVar()
        self.baptised = StringVar()
        self.residence = StringVar()
        self.department = StringVar()
        self.tel = StringVar()
        self.nationality = StringVar()
        self.email = StringVar()
        self.image_path_var = StringVar()
        self.addnav = Frame(self.addwin,width = 1095, height = 80,bg = "#008080")
        self.addnav.place(relx=0,rely = 0)

        self.title= Label(self.addnav,text = "Registration",font =('Montserrat',16,"bold"),bg= "#008080",fg ="white")
        self.title.place(relx = 0.4,rely = 0.4)

        self.image_lbl = Button(self.addnav,bg = '#008080',bd = 0,command=self.back)
        self.logob = Image.open('img/back.png')
        self.resizedb= self.logob.resize((30,30))
        self.real_imageb =ImageTk.PhotoImage(self.resizedb)
        self.image_lbl.config(image =self.real_imageb)
        self.image_lbl.place(relx = 0.95,rely= 0.2)

        self.main = Frame(self.addwin,width=1095,height = 600,bg = 'white')
        self.main.place(relx =0,rely = .11)

        
        self.std_id = Label(self.main,text = "Membership ID",font =('Montserrat',14))
        self.std_id.place(relx = 0.1,rely = 0.1)
        self.std_id_entry1= Entry(self.main,font = ('Montserrat',14),bd = 2,textvariable = self.memID)
        self.std_id_entry1.place(relx = 0.1,rely = 0.16)
        self.gen_id()

        self.std_regno = Label(self.main, text="Full Name", font = ('Montserrat',14))
        self.std_regno.place(relx=0.1, rely=0.24)
        self.std_regno_entry = Entry(self.main, font = ('Montserrat',14),bd = 2,textvariable =self.name)
        self.std_regno_entry.place(relx=0.1, rely=0.3)

        self.std_id = Label(self.main, text="Date of Birth", font = ('Montserrat',14),)
        self.std_id.place(relx=0.1, rely=0.37)
        self.std_id_entry = DateEntry(self.main, font = ('Montserrat',14),textvariable = self.dob,date_pattern = "dd-mm-yyyy",width = 18)
        self.std_id_entry.place(relx=0.1, rely=0.43)

        self.std_id = Label(self.main, text="Gender", font = ('Montserrat',14),)
        self.std_id.place(relx=0.1, rely=0.5)
        self.std_id_entry = ttk.Combobox(self.main, font = ('Montserrat',14),textvariable = self.gender,value= (
        "Male","Female"),width=18)
        self.std_id_entry.place(relx=0.1, rely=0.56)
        
        self.std_id = Label(self.main, text="Are you Baptised", font = ('Montserrat',14),)
        self.std_id.place(relx=0.1, rely=0.63)
        self.std_id_entry = ttk.Combobox(self.main, font = ('Montserrat',14),values=("Yes","No"),width=18,textvariable = self.baptised)
        self.std_id_entry.place(relx=0.1, rely=0.7)

        self.std_id = Label(self.main, text="Residence", font = ('Montserrat',14), )
        self.std_id.place(relx=0.5, rely=0.1)
        self.std_id_entry = Entry(self.main, font = ('Montserrat',14), bd=2,textvariable = self.residence)
        self.std_id_entry.place(relx=0.5, rely=0.16)

        self.std_id = Label(self.main, text="Department", font = ('Montserrat',14), )
        self.std_id.place(relx=0.5, rely=0.24)
        self.std_id_entry = ttk.Combobox(self.main, font = ('Montserrat',14),textvariable = self.department,value= (
        "Evangelism","Choir & Musicians","Children","Finance","Media & Tech. Team","Deacon/Deaconess"),width = 18)
        self.std_id_entry.place(relx=0.5, rely=0.3)

        self.std_id = Label(self.main, text="Phone Number", font = ('Montserrat',14),)
        self.std_id.place(relx=0.5, rely=0.37)
        self.std_id_entry = Entry(self.main, font = ('Montserrat',14), bd=2,textvariable = self.tel)
        self.std_id_entry.place(relx=0.5, rely=0.43)

        self.std_id = Label(self.main, text="Nationality", font = ('Montserrat',14),)
        self.std_id.place(relx=0.5, rely=0.5)
        self.std_id_entry = Entry(self.main, font = ('Montserrat',14), bd=2,textvariable = self.nationality)
        self.std_id_entry.place(relx=0.5, rely=0.56)

        self.std_id = Label(self.main, text="Email Address", font = ('Montserrat',14))
        self.std_id.place(relx=0.5, rely=0.63)
        self.std_id_entry = Entry(self.main, font = ('Montserrat',14), bd=2,textvariable = self.email)
        self.std_id_entry.place(relx=0.5, rely=0.7)

        # Passport Photo
        self.photo = Frame(self.main,width = 200,height=250,bd =1,relief='sunken')
        self.photo.place(relx = .77,rely = .16)
        self.btnUpload =Button(self.main,text='Upload Photo',width =16,font = ('Montserrat',14),bd = 0,bg = "#008080",fg = "white",command = self.upload_photo)
        self.btnUpload.place(relx = .77,rely =.58)

        self.btnUpload =Button(self.main,text='Take Photo',width =16,font = ('Montserrat',14),bd = 0,bg = "#008080",fg = "white",command= self.open_capture_window)
        self.btnUpload.place(relx = .77,rely =.67)

        # sunmit and reset button
        self.submitBtn =Button(self.main,text='Submit Details',width =28,font = ('Montserrat',14),bd = 0,bg = "#008080",fg = "white",command = self.insert_data)
        self.submitBtn.place(relx = .1,rely =.84)
        
        self.submitBtn =Button(self.main,text='Reset Details',width =28,font = ('Montserrat',14),bd = 0,bg = "red",fg = "white",command = self.Reset)
        self.submitBtn.place(relx = .46,rely =.84)
        
    def allMembersData(self):
        self.allwin = Toplevel()
        self.allwin.geometry('1095x680+270+111')
        self.allwin.config(bg ='#008080')
        self.allwin.overrideredirect(True)

        self.search_string = StringVar()
        self.addnav = Frame(self.allwin,width = 1095, height = 80,bg = "#008080")
        self.addnav.place(relx=0,rely = 0)

        self.title= Label(self.addnav,text = "Database",font =('Montserrat',16,"bold"),bg= "#008080",fg ="white")
        self.title.place(relx = 0.4,rely = 0.4)

        self.allback = Button(self.addnav,bg = '#008080',bd = 0,command=self.allwinBack)
        self.logob = Image.open('img/back.png')
        self.resizedb= self.logob.resize((30,30))
        self.real_imageb =ImageTk.PhotoImage(self.resizedb)
        self.allback.config(image =self.real_imageb)
        self.allback.place(relx = 0.95,rely= 0.2)



        self.main = Frame(self.allwin,width=1095,height = 600,bg = 'white')
        self.main.place(relx =0,rely = .11)

        self.printer = Button(self.main,bg = 'white',bd = 0,command=self.generate_pdf)
        self.logop = Image.open('img/printer.png')
        self.resizedp= self.logop.resize((40,30))
        self.real_imagep =ImageTk.PhotoImage(self.resizedp)
        self.printer.config(image =self.real_imagep)
        self.printer.place(relx = 0.87,rely= 0.01)

        self.excel = Button(self.main,bg = 'white',bd = 0,command =self.sheet)
        self.logoe = Image.open('img/excel.png')
        self.resizede= self.logoe.resize((40,30))
        self.real_image_e =ImageTk.PhotoImage(self.resizede)
        self.excel.config(image =self.real_image_e)
        self.excel.place(relx = 0.8,rely= 0.01)

        self.pdf = Button(self.main,bg = 'white',bd = 0,command = self.generate_pdf)
        self.logopdf = Image.open('img/pdf.png')
        self.resizedpdf= self.logopdf.resize((40,30))
        self.real_image_pdf =ImageTk.PhotoImage(self.resizedpdf)
        self.pdf.config(image =self.real_image_pdf)
        self.pdf.place(relx = 0.73,rely= 0.01)

        self.searchBar = Entry(self.main,width= 25,font =('Montserrat',14),bd = 2,relief = SUNKEN,textvariable=self.search_string)
        self.searchBar.place(relx =0.05,rely=0.02)
        

        self.searchbtn = Button(self.main,bg = 'white',bd = 0,command = self.search_data)
        self.searchImage = Image.open('img/search.png')
        self.resizedsearch= self.searchImage.resize((23,23))
        self.real_Search =ImageTk.PhotoImage(self.resizedsearch)
        self.searchbtn.config(image = self.real_Search)
        self.searchbtn.place(relx = 0.37,rely= 0.02)

        self.deletebtn = Button(self.main,bg = 'white',bd = 0,command = self.delete)
        self.deleteImage = Image.open('img/delete.png')
        self.resizeddelete= self.deleteImage.resize((40,30))
        self.real_delete =ImageTk.PhotoImage(self.resizeddelete)
        self.deletebtn.config(image = self.real_delete)
        self.deletebtn.place(relx = 0.66,rely= 0.01)

        self.tree_frame = Frame(self.main,width = 1000,height = 500,bg = "white",bd = 2,relief = SUNKEN)
        self.tree_frame.place(relx = 0.05,rely = 0.1)

        self.scrollbarx = Scrollbar(self.tree_frame, orient=HORIZONTAL)
        self.scrollbary = Scrollbar(self.tree_frame, orient=VERTICAL)
        self.treeview = ttk.Treeview(self.tree_frame, columns=(
            "S/N", "student_id",  "Student Name","Date Of Birth","Place Of Birth", "Name Of Parent","Mobile Number", "class","Residence","Date Of Registration",
        "email"), selectmode="extended", height=15,yscrollcommand=self.scrollbary.set,xscrollcommand=self.scrollbarx.set)
        self.treeview.place(relx=0, rely=0)
        style = ttk.Style()
        # Pick a theme
        style.theme_use('clam')

        style.configure("Treeview.Heading", font = ('Montserrat',14), foreground='#008080',
                        fieldbackground="silver")
        style.configure("Treeview", highlightthickness=4, bd=2, font = ('Montserrat',14), background="#008080",
                        fg="white"
                        , rowheight=40, fieldbackground="silver")
        style.map('Treeview', background=[('selected', 'black')],foreground=[('selected', 'white')])
        self.scrollbary.config(command=self.treeview.yview)
        self.scrollbary.place(relx=0.98, rely=0.01, height=480)
        self.scrollbarx.config(command=self.treeview.xview)
        self.scrollbarx.place(relx=0.01, rely=0.95, width=960)

        self.treeview.heading("S/N", text="S/N", anchor=W)
        self.treeview.heading("student_id", text="Membership ID", anchor=W)
        self.treeview.heading("Student Name", text="Full Name", anchor=W)
        self.treeview.heading("Date Of Birth", text="Gender")
        self.treeview.heading("Place Of Birth", text="Date Of Birth", anchor=W)
        self.treeview.heading("Name Of Parent", text="Confirm Baptism", anchor=W)
        self.treeview.heading("Mobile Number", text="Residence", anchor=W)
        self.treeview.heading("class", text="Department", anchor =W)
        self.treeview.heading("Residence", text="Phone", anchor=W)
        self.treeview.heading("Date Of Registration", text="Nationality", anchor=W)
        self.treeview.heading("email", text="Email", anchor=W)

        self.treeview.column('#0', stretch=NO, minwidth=0, width=0)
        self.treeview.column('#1', stretch=NO, minwidth=0, width=50)
        self.treeview.column('#2', stretch=NO, minwidth=0, width=250)
        self.treeview.column('#3', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#4', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#5', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#6', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#7', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#8', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#9', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#10', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#10', stretch=NO, minwidth=0, width=200)
        self.treeview.tag_configure('oddrow', background='white')
        self.treeview.tag_configure('evenrow', background='#008080')

        self.treeview.place(relx=0., rely=0., width=975, height=470)
        self.treeview.bind("<ButtonRelease-1>", self.on_treeview_click)
        self.DisplayData_members()


    def on_treeview_click(self, event):
        item = self.treeview.selection()
        if item:
            member_id = self.treeview.item(item, 'values')[0]
            self.show_member_details(member_id)

    def show_member_details(self, member_id):
        def dBack():
            self.details_window.destroy()

        # Close the currently open detail window, if any
        if hasattr(self, 'details_window') and self.details_window:
            self.details_window.destroy()

        member_data = membership_backend.view_member_with_image(member_id)
        if member_data:
            self.details_window = Toplevel(self.allwin)
            self.details_window.title("Member Details")
            self.details_window.geometry('500x680+570+111')
            self.details_window.config(bg='white')
            self.details_window.transient(self.allwin)
            self.details_window.iconbitmap('img/icon2.ico')

            self.sback = Button(self.details_window, bg='white', bd=0, command=dBack)
            self.logos = Image.open('img/back.png')
            self.resizeds = self.logos.resize((40, 40))
            self.real_images = ImageTk.PhotoImage(self.resizeds)
            self.sback.config(image=self.real_images)
            self.sback.place(relx=0.4, rely=0.85)

            labels = ["ID", "MemID", "Name", "Gender", "DOB", "Baptised", "Residence", "Department", "Tel", "Nationality", "Email"]
            for label, value in zip(labels, member_data[:-1]):
                Label(self.details_window, text=f"{label}:", font=('Montserrat', 14), bg="white").place(relx=0.05, rely=0.3 + labels.index(label) * 0.05, anchor='w')
                Label(self.details_window, text=value, font=('Montserrat', 14), bg="white", fg="black").place(relx=0.5, rely=0.3 + labels.index(label) * 0.05, anchor='w')

            image_data = member_data[-1]
            if image_data:
                image = Image.open(io.BytesIO(image_data))
                image = image.resize((130, 150), Image.LANCZOS)
                photo = ImageTk.PhotoImage(image)
                photo_Frame = Frame(self.details_window, width=130, height=150, bd=2, relief=SUNKEN)
                photo_Frame.place(relx=.3, rely=0)
                img_label = Label(photo_Frame, image=photo)
                img_label.image = photo
                img_label.place(relx=0, rely=0)

                # Important: Keep a reference to the image to prevent garbage collection
                self.details_window.image_reference = photo
            else:
                photo_Frame = Frame(self.details_window, width=130, height=150, bd=2, relief=SUNKEN)
                photo_Frame.place(relx=.3, rely=0)
                Label(photo_Frame, text="No image", pady=10, font=('Montserrat', 12), fg="red",).place(relx=.1, rely=0.4)

            

    def Giving(self):
        self.givewin = Toplevel()
        self.givewin.geometry('1095x680+270+111')
        self.givewin.config(bg ='#008080')
        self.givewin.overrideredirect(True)


        self.addnav = Frame(self.givewin,width = 1095, height = 80,bg = "#008080")
        self.addnav.place(relx=0,rely = 0)

        self.title= Label(self.addnav,text = "Giving",font =('Montserrat',16,"bold"),bg= "#008080",fg ="white")
        self.title.place(relx = 0.4,rely = 0.4)

        self.image_lbl = Button(self.addnav,bg = '#008080',bd = 0,command=self.giveBack)
        self.logob = Image.open('img/back.png')
        self.resizedb= self.logob.resize((30,30))
        self.real_imageb =ImageTk.PhotoImage(self.resizedb)
        self.image_lbl.config(image =self.real_imageb)
        self.image_lbl.place(relx = 0.95,rely= 0.2)

        self.main = Frame(self.givewin,width=1095,height = 600,bg = 'white')
        self.main.place(relx =0,rely = .11)
        self.tithFrame()
    
    def tithFrame(self):

        self.member_ID = StringVar()
        self.name = StringVar()
        self.date = StringVar()
        self.amount = StringVar()
        self.contact = StringVar()

        # giving variables
        self.gtype = StringVar()
        self.offer_amt = StringVar()
        self.offer_date = StringVar()
        self.service = StringVar()


        self.dataFrame = LabelFrame(self.main,text = "Tithe",width = 650, height = 400, bg = 'white',bd = 2,relief=RAISED)
        self.dataFrame.place(relx = .05, rely = .04)
        self.msgID = Label(self.dataFrame,text = 'Membership ID',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .05)
        self.msgID_entry = Entry(self.dataFrame,width = 28,font =('Montserrat',14),bg= "white",textvariable=self.member_ID) 
        self.msgID_entry.place(relx=.3,rely = .05)
        self.gen_id_tithe()

        self.msgID = Label(self.dataFrame,text = 'Full Name',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .15)
        self.msgID_entry = Entry(self.dataFrame,width = 28,font =('Montserrat',14),bg= "white",textvariable=self.name) 
        self.msgID_entry.place(relx=.3,rely = .15)

        self.msgID = Label(self.dataFrame,text = 'Date of Payment',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .25)
        self.std_id_entry = DateEntry(self.dataFrame, font = ('Montserrat',14),date_pattern = "dd-mm-yyyy",width = 27,textvariable = self.date)
        self.std_id_entry.place(relx=0.3, rely=0.25)

        self.msgID = Label(self.dataFrame,text = 'Amount Paid',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .35)
        self.msgID_entry = Entry(self.dataFrame,width = 28,font =('Montserrat',14),bg= "white",textvariable = self.amount) 
        self.msgID_entry.place(relx=.3,rely = .35)

        self.msgID = Label(self.dataFrame,text = 'Contact',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .45)
        self.msgID_entry = Entry(self.dataFrame,width = 28,font =('Montserrat',14),bg= "white",textvariable = self.contact) 
        self.msgID_entry.place(relx=.3,rely = .45)

        self.sendmsg = Button(self.dataFrame,text = 'Save Details',font =('Montserrat',14,'bold'),bg ="#008080",fg = 'white',width = 20,command = self.InsertTithe)
        self.sendmsg.place(relx = 0.3,rely = .63)

        # self.sendmsg = Button(self.dataFrame,text = 'Send Receipt',font =('Montserrat',14,'bold'),bg ="#008080",fg = 'white',width = 20,command = 'self.Tithe_sms')
        # self.sendmsg.place(relx = 0.5,rely = .63)
        self.resetmsg = Button(self.main,text = 'Reset',font =('Montserrat',14,'bold'),bg ="red",fg = 'white',width = 75)
        self.resetmsg.place(relx = 0.05,rely = .76)

        self.resetmsg = Button(self.dataFrame,text = 'Show Database',font =('Montserrat',14,'bold'),bg ="#008080",fg = 'white',width = 40,command =  self.TitheData)
        self.resetmsg.place(relx = 0.1,rely = .8)

        self.dataFrame2 = LabelFrame(self.main,text = "Other offerings",width = 330, height = 400, bg = 'white',bd = 2,relief='sunken')
        self.dataFrame2.place(relx = .65, rely = .04)

        self.offer_type = Label(self.dataFrame2,text = "Type",font =('Montserrat',14),bg= "white")
        self.offer_type.place(relx =.03,rely =.05)
        self.std_id_entry = ttk.Combobox(self.dataFrame2, font = ('Montserrat',14),value= (
        "Mission Funds","Boosters","Pledge"),width = 12,textvariable=self.gtype)
        self.std_id_entry.place(relx=0.3, rely=0.05)

        self.msgID = Label(self.dataFrame2,text = 'Amount',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .15)
        self.msgID_entry = Entry(self.dataFrame2,width = 14,font =('Montserrat',14),bg= "white",textvariable=self.offer_amt) 
        self.msgID_entry.place(relx=.3,rely = .15)

        self.msgID = Label(self.dataFrame2,text = 'Date',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .25)
        self.std_id_entry = DateEntry(self.dataFrame2, font = ('Montserrat',14),date_pattern = "dd-mm-yyyy",width = 12,textvariable = self.offer_date)
        self.std_id_entry.place(relx=0.3, rely=0.25)

        self.msgID = Label(self.dataFrame2,text = 'Service',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .35)
        self.msgID_entry = Entry(self.dataFrame2,width = 14,font =('Montserrat',14),bg= "white",textvariable = self.service) 
        self.msgID_entry.place(relx=.3,rely = .35)

        self.save_offer = Button(self.dataFrame2,text = 'Save Data',font =('Montserrat',14,'bold'),bg ="#008080",fg = 'white',width = 20, command = self.InsertOffer)
        self.save_offer.place(relx = 0.1,rely = .63)

        self.showmsg = Button(self.dataFrame2,text = 'Show Database',font =('Montserrat',14,'bold'),bg ="#008080",fg = 'white',width = 20,command = self.GiveData)
        self.showmsg.place(relx = 0.1,rely = .8)

    def InsertTithe(self):
        id_num = self.member_ID.get()
        name = self.name.get()
        date = self.date.get()
        amt = self.amount.get()
        contact = self.contact.get()

        if(id_num and amt and date and contact !=" "):
            giving_backend.TitheInsert(id_num,name,date,amt,contact)
            self.Tithe_sms()
            messagebox.showinfo("sucess","Data Stored Successfully",parent = self.givewin)
            
        else:
            messagebox.showerror("error","All fields are required",parent = self.givewin)
    def Tithe_sms(self):
        api_url = "https://sms.arkesel.com/sms/api?action=send-sms"
        api_key = "OjBxSFBoQ1NrUFJ6Q0MwR0s="

        phone_number = self.contact.get
        ()
        tithername = self.name.get()
        titherID = self.member_ID.get()
        sender_id = 'Tithe-MSG'
        date = self.date.get()
        amount = self.amount.get()
        message = f"""
        Dear {tithername}({titherID}),you have paid an amount of GHS{amount} as Tithe on {date}. God bless you for being a covenant tither.
        """
        if not (phone_number and titherID and tithername):
            messagebox.showwarning("Warning", "Please fill out all fields.", parent=self.givewin)
            return

        try:
            response = requests.get(f"{api_url}&api_key={api_key}&to={phone_number}&from={sender_id}&sms={message}")

            if response.status_code == 200:
                messagebox.showinfo("Success", f"SMS sent successfully!", parent=self.givewin)
            else:
                messagebox.showerror("Error", f"Failed to send SMS. Status Code: {response.status_code}", parent=self.givewin)
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Error", f"Error sending SMS: {e}", parent=self.givewin)

    def InsertOffer(self):
        gtype = self.gtype.get()
        gamt = self.offer_amt.get()
        gdate = self.offer_date.get()
        gservice = self.service.get()

        if(gtype and gamt and gdate and gservice !=" "):
            giving_backend.GiveInsert(gtype,gamt,gdate,gservice)
            messagebox.showinfo("sucess","Data Stored Successfully",parent = self.givewin)
        else:
            messagebox.showerror("error","All fields are required",parent = self.givewin)

    def TitheData(self):
        self.search_string_tithe = StringVar()
        self.main = Frame(self.givewin,width=1095,height = 600,bg = 'white')
        self.main.place(relx =0,rely = .11)

        self.printer = Button(self.main,bg = 'white',bd = 0,command = self.generate_pdf_tithe)
        self.logop = Image.open('img/printer.png')
        self.resizedp= self.logop.resize((40,30))
        self.real_imagep =ImageTk.PhotoImage(self.resizedp)
        self.printer.config(image =self.real_imagep)
        self.printer.place(relx = 0.75,rely= 0.01)

        self.closeBtn=Button(self.main,text = "Close",font =('Montserrat',13),bg = "red",fg = "white",bd = 0,width = 13,command = self.CloseTithDB)
        self.closeBtn.place(relx = 0.84,rely =0.01)

        self.excel = Button(self.main,bg = 'white',bd = 0,command = self.sheet_Tithe)
        self.logoe = Image.open('img/excel.png')
        self.resizede= self.logoe.resize((40,30))
        self.real_image_e =ImageTk.PhotoImage(self.resizede)
        self.excel.config(image =self.real_image_e)
        self.excel.place(relx = 0.67,rely= 0.01)

        self.pdf = Button(self.main,bg = 'white',bd = 0,command = self.generate_pdf)
        self.logopdf = Image.open('img/pdf.png')
        self.resizedpdf= self.logopdf.resize((40,30))
        self.real_image_pdf =ImageTk.PhotoImage(self.resizedpdf)
        self.pdf.config(image =self.real_image_pdf)
        self.pdf.place(relx = 0.6,rely= 0.01)

        self.searchBar = Entry(self.main,width= 25,font =('Montserrat',14),bd = 2,relief = SUNKEN,textvariable=self.search_string_tithe)
        self.searchBar.place(relx =0.05,rely=0.02)

        self.searchbtn = Button(self.main,bg = 'white',bd = 0,command = self.search_data_Tithe)
        self.searchImage = Image.open('img/search.png')
        self.resizedsearch= self.searchImage.resize((23,23))
        self.real_Search =ImageTk.PhotoImage(self.resizedsearch)
        self.searchbtn.config(image = self.real_Search)
        self.searchbtn.place(relx = 0.37,rely= 0.02)

        self.deletebtn = Button(self.main,bg = 'white',bd = 0,command = self.delete_tithe)
        self.deleteImage = Image.open('img/delete.png')
        self.resizeddelete= self.deleteImage.resize((40,30))
        self.real_delete =ImageTk.PhotoImage(self.resizeddelete)
        self.deletebtn.config(image = self.real_delete)
        self.deletebtn.place(relx = 0.54,rely= 0.01)

        self.tree_frame = Frame(self.main,width = 1000,height = 500,bg = "white",bd = 2,relief = SUNKEN)
        self.tree_frame.place(relx = 0.05,rely = 0.1)

        self.scrollbarx = Scrollbar(self.tree_frame, orient=HORIZONTAL)
        self.scrollbary = Scrollbar(self.tree_frame, orient=VERTICAL)
        self.treeview = ttk.Treeview(self.tree_frame, columns=(
            "S/N", "mem_id",  "Full Name","Date Of Payment","Amount Paid", "Contact",
        ), selectmode="extended", height=15,yscrollcommand=self.scrollbary.set,xscrollcommand=self.scrollbarx.set)
        self.treeview.place(relx=0, rely=0)
        style = ttk.Style()
        # Pick a theme
        style.theme_use('clam')

        style.configure("Treeview.Heading", font = ('Montserrat',14), foreground='#008080',
                        fieldbackground="silver")
        style.configure("Treeview", highlightthickness=4, bd=2, font = ('Montserrat',14), background="#008080",
                        fg="white"
                        , rowheight=40, fieldbackground="silver")
        style.map('Treeview', background=[('selected', 'black')],foreground=[('selected', 'white')])
        self.scrollbary.config(command=self.treeview.yview)
        self.scrollbary.place(relx=0.98, rely=0.01, height=480)
        self.scrollbarx.config(command=self.treeview.xview)
        self.scrollbarx.place(relx=0.01, rely=0.95, width=960)

        self.treeview.heading("S/N", text="S/N", anchor=W)
        self.treeview.heading("mem_id", text="Membership ID", anchor=W)
        self.treeview.heading("Full Name", text="Full Name", anchor=W)
        self.treeview.heading("Date Of Payment", text="Date Of Payment")
        self.treeview.heading("Amount Paid", text="Amount Paid", anchor=W)
        self.treeview.heading("Contact", text="Contact", anchor=W)
      

        self.treeview.column('#0', stretch=NO, minwidth=0, width=0)
        self.treeview.column('#1', stretch=NO, minwidth=0, width=50)
        self.treeview.column('#2', stretch=NO, minwidth=0, width=250)
        self.treeview.column('#3', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#4', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#5', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#6', stretch=NO, minwidth=0, width=200)
        self.treeview.tag_configure('oddrow', background='white')
        self.treeview.tag_configure('evenrow', background='#008080')

        self.treeview.place(relx=0., rely=0., width=975, height=470)

        self.DisplayData_Tithe()

    def GiveData(self):
        self.search_string_give = StringVar()
        self.main = Frame(self.givewin,width=1095,height = 600,bg = 'white')
        self.main.place(relx =0,rely = .11)

        self.printer = Button(self.main,bg = 'white',bd = 0,command = self.generate_pdf_give)
        self.logop = Image.open('img/printer.png')
        self.resizedp= self.logop.resize((40,30))
        self.real_imagep =ImageTk.PhotoImage(self.resizedp)
        self.printer.config(image =self.real_imagep)
        self.printer.place(relx = 0.75,rely= 0.01)

        self.closeBtn=Button(self.main,text = "Close",font =('Montserrat',13),bg = "red",fg = "white",bd = 0,width = 13,command = self.CloseTithDB)
        self.closeBtn.place(relx = 0.84,rely =0.01)

        self.excel = Button(self.main,bg = 'white',bd = 0,command = self.sheet_Give)
        self.logoe = Image.open('img/excel.png')
        self.resizede= self.logoe.resize((40,30))
        self.real_image_e =ImageTk.PhotoImage(self.resizede)
        self.excel.config(image =self.real_image_e)
        self.excel.place(relx = 0.67,rely= 0.01)

        self.pdf = Button(self.main,bg = 'white',bd = 0,command = self.generate_pdf)
        self.logopdf = Image.open('img/pdf.png')
        self.resizedpdf= self.logopdf.resize((40,30))
        self.real_image_pdf =ImageTk.PhotoImage(self.resizedpdf)
        self.pdf.config(image =self.real_image_pdf)
        self.pdf.place(relx = 0.6,rely= 0.01)

        self.searchBar = Entry(self.main,width= 25,font =('Montserrat',14),bd = 2,relief = SUNKEN,textvariable=self.search_string_give)
        self.searchBar.place(relx =0.05,rely=0.02)

        self.searchbtn = Button(self.main,bg = 'white',bd = 0,command = self.search_data_Give)
        self.searchImage = Image.open('img/search.png')
        self.resizedsearch= self.searchImage.resize((23,23))
        self.real_Search =ImageTk.PhotoImage(self.resizedsearch)
        self.searchbtn.config(image = self.real_Search)
        self.searchbtn.place(relx = 0.37,rely= 0.02)

        self.deletebtn = Button(self.main,bg = 'white',bd = 0,command = self.delete_Give)
        self.deleteImage = Image.open('img/delete.png')
        self.resizeddelete= self.deleteImage.resize((40,30))
        self.real_delete =ImageTk.PhotoImage(self.resizeddelete)
        self.deletebtn.config(image = self.real_delete)
        self.deletebtn.place(relx = 0.54,rely= 0.01)

        self.tree_frame = Frame(self.main,width = 1000,height = 500,bg = "white",bd = 2,relief = SUNKEN)
        self.tree_frame.place(relx = 0.05,rely = 0.1)

        self.scrollbarx = Scrollbar(self.tree_frame, orient=HORIZONTAL)
        self.scrollbary = Scrollbar(self.tree_frame, orient=VERTICAL)
        self.treeview = ttk.Treeview(self.tree_frame, columns=(
            "S/N", "Type of Offering",  "Amount","Date Of Payment","Service"
        ), selectmode="extended", height=15,yscrollcommand=self.scrollbary.set,xscrollcommand=self.scrollbarx.set)
        self.treeview.place(relx=0, rely=0)
        style = ttk.Style()
        # Pick a theme
        style.theme_use('clam')

        style.configure("Treeview.Heading", font = ('Montserrat',14), foreground='#008080',
                        fieldbackground="silver")
        style.configure("Treeview", highlightthickness=4, bd=2, font = ('Montserrat',14), background="#008080",
                        fg="white"
                        , rowheight=40, fieldbackground="silver")
        style.map('Treeview', background=[('selected', 'black')],foreground=[('selected', 'white')])
        self.scrollbary.config(command=self.treeview.yview)
        self.scrollbary.place(relx=0.98, rely=0.01, height=480)
        self.scrollbarx.config(command=self.treeview.xview)
        self.scrollbarx.place(relx=0.01, rely=0.95, width=960)

        self.treeview.heading("S/N", text="S/N", anchor=W)
        self.treeview.heading("Type of Offering", text="Type of Offering", anchor=W)
        self.treeview.heading("Amount", text="Amount", anchor=W)
        self.treeview.heading("Date Of Payment", text="Date Of Payment")
        self.treeview.heading("Service", text="Service", anchor=W)
      
        self.treeview.column('#0', stretch=NO, minwidth=0, width=0)
        self.treeview.column('#1', stretch=NO, minwidth=0, width=50)
        self.treeview.column('#2', stretch=NO, minwidth=0, width=250)
        self.treeview.column('#3', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#4', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#5', stretch=NO, minwidth=0, width=300)
        
        self.treeview.tag_configure('oddrow', background='white')
        self.treeview.tag_configure('evenrow', background='#008080')

        self.treeview.place(relx=0., rely=0., width=975, height=470)

        self.DisplayData_Give()

    # Attendance Data
    def attendanceData(self):
        self.main = Frame(self.attendwin,width=1095,height = 600,bg = 'white')
        self.main.place(relx =0,rely = .11)

        self.printer = Button(self.main,bg = 'white',bd = 0,command = self.generate_pdf_give)
        self.logop = Image.open('img/printer.png')
        self.resizedp= self.logop.resize((40,30))
        self.real_imagep =ImageTk.PhotoImage(self.resizedp)
        self.printer.config(image =self.real_imagep)
        self.printer.place(relx = 0.75,rely= 0.01)

        self.closeBtn=Button(self.main,text = "Close",font =('Montserrat',13),bg = "red",fg = "white",bd = 0,width = 13,command = self.CloseTithDB)
        self.closeBtn.place(relx = 0.84,rely =0.01)

        self.excel = Button(self.main,bg = 'white',bd = 0,command = self.sheet_attend)
        self.logoe = Image.open('img/excel.png')
        self.resizede= self.logoe.resize((40,30))
        self.real_image_e =ImageTk.PhotoImage(self.resizede)
        self.excel.config(image =self.real_image_e)
        self.excel.place(relx = 0.67,rely= 0.01)

        self.pdf = Button(self.main,bg = 'white',bd = 0,command = self.generate_pdf)
        self.logopdf = Image.open('img/pdf.png')
        self.resizedpdf= self.logopdf.resize((40,30))
        self.real_image_pdf =ImageTk.PhotoImage(self.resizedpdf)
        self.pdf.config(image =self.real_image_pdf)
        self.pdf.place(relx = 0.6,rely= 0.01)

        self.searchBar = Entry(self.main,width= 25,font =('Montserrat',14),bd = 2,relief = SUNKEN,textvariable=self.search_string_attend)
        self.searchBar.place(relx =0.05,rely=0.02)

        self.searchbtn = Button(self.main,bg = 'white',bd = 0,command = self.search_attendace)
        self.searchImage = Image.open('img/search.png')
        self.resizedsearch= self.searchImage.resize((23,23))
        self.real_Search =ImageTk.PhotoImage(self.resizedsearch)
        self.searchbtn.config(image = self.real_Search)
        self.searchbtn.place(relx = 0.37,rely= 0.02)

        self.deletebtn = Button(self.main,bg = 'white',bd = 0,command = self.delete_attend)
        self.deleteImage = Image.open('img/delete.png')
        self.resizeddelete= self.deleteImage.resize((40,30))
        self.real_delete =ImageTk.PhotoImage(self.resizeddelete)
        self.deletebtn.config(image = self.real_delete)
        self.deletebtn.place(relx = 0.54,rely= 0.01)

        self.tree_frame = Frame(self.main,width = 1000,height = 500,bg = "white",bd = 2,relief = SUNKEN)
        self.tree_frame.place(relx = 0.05,rely = 0.1)

        self.scrollbarx = Scrollbar(self.tree_frame, orient=HORIZONTAL)
        self.scrollbary = Scrollbar(self.tree_frame, orient=VERTICAL)
        self.treeview = ttk.Treeview(self.tree_frame, columns=(
            "S/N", "date",  "program","male","female","youth","Children","total"
        ), selectmode="extended", height=15,yscrollcommand=self.scrollbary.set,xscrollcommand=self.scrollbarx.set)
        self.treeview.place(relx=0, rely=0)
        style = ttk.Style()
        # Pick a theme
        style.theme_use('clam')

        style.configure("Treeview.Heading", font = ('Montserrat',14), foreground='#008080',
                        fieldbackground="silver")
        style.configure("Treeview", highlightthickness=4, bd=2, font = ('Montserrat',14), background="#008080",
                        fg="white"
                        , rowheight=40, fieldbackground="silver")
        style.map('Treeview', background=[('selected', 'black')],foreground=[('selected', 'white')])
        self.scrollbary.config(command=self.treeview.yview)
        self.scrollbary.place(relx=0.98, rely=0.01, height=480)
        self.scrollbarx.config(command=self.treeview.xview)
        self.scrollbarx.place(relx=0.01, rely=0.95, width=960)

        self.treeview.heading("S/N", text="S/N", anchor=W)
        self.treeview.heading("date", text="Date", anchor=W)
        self.treeview.heading("program", text="Program/Service", anchor=W)
        self.treeview.heading("male", text="Total Number of Men")
        self.treeview.heading("female", text="Total Number of Women", anchor=W)
        self.treeview.heading("youth", text="Total Number of Youth")
        self.treeview.heading("Children", text="Total Number of Children", anchor=W)
        self.treeview.heading("total", text="Total Attendance", anchor=W)
      

        self.treeview.column('#0', stretch=NO, minwidth=0, width=0)
        self.treeview.column('#1', stretch=NO, minwidth=0, width=50)
        self.treeview.column('#2', stretch=NO, minwidth=0, width=200)
        self.treeview.column('#3', stretch=NO, minwidth=0, width=250)
        self.treeview.column('#4', stretch=NO, minwidth=0, width=250)
        self.treeview.column('#5', stretch=NO, minwidth=0, width=250)
        self.treeview.column('#6', stretch=NO, minwidth=0, width=250)
        self.treeview.column('#7', stretch=NO, minwidth=0, width=250)
        
        self.treeview.tag_configure('oddrow', background='white')
        self.treeview.tag_configure('evenrow', background='#008080')

        self.treeview.place(relx=0., rely=0., width=975, height=470)

        self.DisplayData_attend()

    def CloseTithDB(self):
        self.main.destroy()



    def instantMessage(self):
        self.msgwin = Toplevel()
        self.msgwin.geometry('1095x680+270+111')
        self.msgwin.config(bg ='#008080')
        self.msgwin.overrideredirect(True)

        self.id = StringVar()
        self.msgline =StringVar()
       

        self.addnav = Frame(self.msgwin,width = 1095, height = 80,bg = "#008080")
        self.addnav.place(relx=0,rely = 0)

        self.title= Label(self.addnav,text = "Messaging",font =('Montserrat',16,"bold"),bg= "#008080",fg ="white")
        self.title.place(relx = 0.4,rely = 0.4)

        self.image_lbl = Button(self.addnav,bg = '#008080',bd = 0,command=self.msgBack)
        self.logob = Image.open('img/back.png')
        self.resizedb= self.logob.resize((30,30))
        self.real_imageb =ImageTk.PhotoImage(self.resizedb)
        self.image_lbl.config(image =self.real_imageb)
        self.image_lbl.place(relx = 0.95,rely= 0.2)

        self.main = Frame(self.msgwin,width=1095,height = 600,bg = 'white')
        self.main.place(relx =0,rely = .11)
        
        self.messagingFrame = Frame(self.main,width = 750, height = 500, bg = 'white',bd = 2,relief='sunken')
        self.messagingFrame.place(relx = .15, rely = .04)

        # details
        self.id.set("HGWC")
        self.msgID = Label(self.messagingFrame,text = 'Message ID',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .05)
        self.msgID_entry = Entry(self.messagingFrame,width = 38,font =('Montserrat',14),bg= "white",textvariable=self.id) 
        self.msgID_entry.place(relx=.2,rely = .05)

        self.reciever = Label(self.messagingFrame,text = 'Contact',font =('Montserrat',14),bg= "white",)
        self.reciever.place(relx = .03,rely = .15)
        self.reciever_entry = Entry(self.messagingFrame,width = 38,font =('Montserrat',14),bg= "white",textvariable=self.msgline) 
        self.reciever_entry.place(relx=.2,rely = .15)

        self.contact = Button(self.messagingFrame,bg = 'white',bd = 0,command = self.load_contacts_from_excel)
        self.logocon = Image.open('img/contact.png')
        self.resizedcon= self.logocon.resize((30,30))
        self.real_imagecon =ImageTk.PhotoImage(self.resizedcon)
        self.contact.config(image =self.real_imagecon)
        self.contact.place(relx = 0.89,rely= 0.15)

        self.body = Label(self.messagingFrame,text = 'Body',font =('Montserrat',14),bg= "white",)
        self.body.place(relx = .03,rely = .25)
        self.body_entry = Text(self.messagingFrame,width =38,height = 9,bd = 2,font =('Montserrat',14))
        self.body_entry.place(relx = .2,rely =.25)

        # Buttons
        self.sendmsg = Button(self.messagingFrame,text = 'Send',font =('Montserrat',14,'bold'),bg ="#008080",fg = 'white',width = 20,command = self.bulkSMS)
        self.sendmsg.place(relx = 0.18,rely = .8)

        self.resetmsg = Button(self.messagingFrame,text = 'Reset',font =('Montserrat',14,'bold'),bg ="red",fg = 'white',width = 20)
        self.resetmsg.place(relx = 0.53,rely = .8)


    
    
    def attendance(self):
        self.attendwin = Toplevel()
        self.attendwin.geometry('1095x680+270+111')
        self.attendwin.config(bg ='#008080')
        self.attendwin.overrideredirect(True)
    

        self.addnav = Frame(self.attendwin,width = 1095, height = 80,bg = "#008080")
        self.addnav.place(relx=0,rely = 0)

        self.title= Label(self.addnav,text = "Attendance",font =('Montserrat',16,"bold"),bg= "#008080",fg ="white")
        self.title.place(relx = 0.4,rely = 0.4)

        self.image_lbl = Button(self.addnav,bg = '#008080',bd = 0,command=self.attendBack)
        self.logob = Image.open('img/back.png')
        self.resizedb= self.logob.resize((30,30))
        self.real_imageb =ImageTk.PhotoImage(self.resizedb)
        self.image_lbl.config(image =self.real_imageb)
        self.image_lbl.place(relx = 0.95,rely= 0.2)

        self.visitors_icon = Button(self.addnav,text = "visitors",font =('Montserrat',12,"bold"),bg= "#008080",fg ="white",command= self.visitors)
        self.visitors_icon.place(relx = 0.8,rely = 0.2)
        self.main = Frame(self.attendwin,width=1095,height = 600,bg = 'white')
        self.main.place(relx =0,rely = .11)

        self.dataFrame3 = Frame(self.main,width = 750, height = 400, bg = 'white',bd = 2,relief=RAISED)
        self.dataFrame3.place(relx = .15, rely = .1)


        self.msgID = Label(self.dataFrame3,text = 'Date of Attendance',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .05)
        self.std_id_entry = DateEntry(self.dataFrame3, font = ('Montserrat',14),date_pattern = "dd-mm-yyyy",width = 27, textvariable = self.attend_date)
        self.std_id_entry.place(relx=0.3, rely=0.05)

        
        self.msgID = Label(self.dataFrame3,text = 'Program',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .15)
        self.msgID_entry = Entry(self.dataFrame3,width = 28,font =('Montserrat',14),bg= "white",textvariable=self.attend_program) 
        self.msgID_entry.place(relx=.3,rely = .15)

        self.msgID = Label(self.dataFrame3,text = 'Number Of Men',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .25)
        self.males_entry = Entry(self.dataFrame3,width = 28,font =('Montserrat',14),bg= "white",textvariable=self.attend_male) 
        self.males_entry.place(relx=.3,rely = .25)
        self.males_entry.bind("<KeyRelease>", self.update_total)


        self.msgID = Label(self.dataFrame3,text = 'Number Of Women',font =('Montserrat',14),bg= "white",)
        self.msgID.place(relx = .03,rely = .35)
        self.females_entry = Entry(self.dataFrame3,width = 28,font =('Montserrat',14),bg= "white",textvariable=self.attend_female) 
        self.females_entry.place(relx=.3,rely = .35)
        self.females_entry.bind("<KeyRelease>", self.update_total)

        self.msgID = Label(self.dataFrame3,text = 'Number Of Youth',font =('Montserrat',14),bg= "white")
        self.msgID.place(relx = .03,rely = .45)
        self.youth_entry = Entry(self.dataFrame3,width = 28,font =('Montserrat',14),bg= "white",textvariable=self.attend_youth) 
        self.youth_entry.place(relx=.3,rely = .45)
        self.youth_entry.bind("<KeyRelease>", self.update_total)

        self.msgID = Label(self.dataFrame3,text = 'Number Of Children',font =('Montserrat',14),bg= "white")
        self.msgID.place(relx = .03,rely = .55)
        self.child_entry = Entry(self.dataFrame3,width = 28,font =('Montserrat',14),bg= "white",textvariable=self.attend_child) 
        self.child_entry.place(relx=.3,rely = .55)
        self.child_entry.bind("<KeyRelease>", self.update_total)

        self.msgID = Label(self.dataFrame3,text = 'Total Attendance',font =('Montserrat',14),bg= "white")
        self.msgID.place(relx = .03,rely = .65)
        self.tot_entry = Entry(self.dataFrame3,width = 28,font =('Montserrat',14),bg= "white",state = "disabled",textvariable=self.attend_total) 
        self.tot_entry.place(relx=.3,rely = .65)

        self.resetmsg = Button(self.dataFrame3,text = 'Save Data',font =('Montserrat',14,'bold'),bg ="#008080",fg = 'white',width = 28,command = self.insert_attendance)
        self.resetmsg.place(relx = 0.02,rely = .8)

        self.resetmsg = Button(self.dataFrame3,text = 'Attendance Database',font =('Montserrat',14,'bold'),bg ="#008080",fg = 'white',width = 28,command = self.attendanceData)
        self.resetmsg.place(relx = 0.46,rely = .8)
    
    def update_total(self,event):
        a = int(self.attend_male.get())
        b = int(self.attend_female.get())
        c = int(self.attend_youth.get())
        d = int(self.attend_child.get())
        
        tot = a+b+c+d
        self.attend_total.set(tot)
        
    def gen_id(self):
        # self.std_id_entry1.config(state='disabled', disabledbackground="#008080", disabledforeground="white")
        self.rand_id = random.randint(100, 900)
        self.conv_id = ('HGWC/GR/AS/24/' + str(self.rand_id))
        self.memID.set(self.conv_id)

    def gen_id_tithe(self):
        self.rand_id = random.randint(100, 900)
        self.conv_id = ('HGWC/GR/AS/24/' + str(self.rand_id))
        self.member_ID.set(self.conv_id)
        
        
    def Developers(self):
        self.profile = "https://www.linkedin.com/in/samuel-kwabena-nyonator-1703562b1"
        self.url = self.profile
        try:
            webbrowser.open(self.url)
        except:
            pass
    def Expenses(self):
        pass
    
    def calendar(self):
        messagebox.showinfo("Developer","This functionality is under development")

    def AiFun(self):
        messagebox.showinfo("Developer","This functionality is under development")

    def TimeDisplay(self):
        self.current_date = datetime.now()
        self.formatted_date = self.current_date.strftime("%A, %d %B, %Y")

        self.datelabel = Label(self.navigation_bar,text=self.formatted_date,font = ('Montserrat',12),bg = "white",fg="black")
        self.datelabel.place(relx=.43,rely = .4)

    def back(self):
        self.addwin.destroy()
    
    def allwinBack(self):
        self.allwin.destroy()
    
    def msgBack(self):
        self.msgwin.destroy()
    
    def giveBack(self):
        self.givewin.destroy()

    def attendBack(self):
        self.attendwin.destroy()

    def voiceBack(self):
        self.voicewin.destroy()

    def upload_photo(self):
        file_path = filedialog.askopenfilename(title="Select an image file",
                                                filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.gif")],
                                                parent=self.addwin)
        
        if file_path:
            
            try:
                image = Image.open(file_path)
                image = image.resize((200, 250), Image.LANCZOS)
                photo = ImageTk.PhotoImage(image)

                existing_label = getattr(self, "label", None)
                if existing_label:
                    existing_label.destroy()

                label = Label(self.photo, image=photo)
                label.photo = photo 
                label.pack()

                self.image_path_var.set(file_path)

                self.label = label

                messagebox.showinfo("File Path",f"Successfully Selected file:{file_path}",parent = self.addwin)
            except Exception as e:
                messagebox.showerror("Error",f"Error opening image: {e}",parent = self.addwin)
    

    def insert_attendance(self):
        # Insert attendance data into the database
        date = self.attend_date.get()
        program = self.attend_program.get()
        male = self.attend_male.get()
        female = self.attend_female.get()
        youth = self.attend_youth.get()
        child = self.attend_child.get()
        total = self.attend_total.get()
        if(date and program and male and female and youth and child !=" "):
            attendace_backend.insert_attend(date,program,male,female,youth,child,total)
            messagebox.showinfo("sucess","Data Stored Successfully",parent = self.attendwin)
        else:
            messagebox.showerror("error","All fields are required",parent = self.attendwin)


    def DisplayData_attend(self):
            self.treeview.delete(*self.treeview.get_children())
            conn = sqlite3.connect('attendance_database.db')
            cursor = conn.execute("SELECT * FROM attendance")
            fetch = cursor.fetchall()
            count = 0
            for data in fetch:
                if count % 2 == 0:
                    self.treeview.insert('', 'end', values=(data), tags=('evenrow'))
                else:
                    self.treeview.insert('', 'end', values=(data), tags=('oddrow'))
                count += 1

            cursor.close()
            conn.close()

    def DisplayData_members(self):
            self.treeview.delete(*self.treeview.get_children())
            conn = sqlite3.connect('database.db')
            cursor = conn.execute("SELECT * FROM members")
            fetch = cursor.fetchall()
            count = 0
            for data in fetch:
                if count % 2 == 0:
                    self.treeview.insert('', 'end', values=(data), tags=('evenrow'))
                else:
                    self.treeview.insert('', 'end', values=(data), tags=('oddrow'))
                count += 1

            cursor.close()
            conn.close()

    def DisplayData_Give(self):
            self.treeview.delete(*self.treeview.get_children())
            conn = sqlite3.connect('GiveDB.db')
            cursor = conn.execute("SELECT * FROM giving")
            fetch = cursor.fetchall()
            count = 0
            for data in fetch:
                if count % 2 == 0:
                    self.treeview.insert('', 'end', values=(data), tags=('evenrow'))
                else:
                    self.treeview.insert('', 'end', values=(data), tags=('oddrow'))
                count += 1

            cursor.close()
            conn.close()

    
    def insert_data(self):
        memID = self.memID.get()
        name = self.name.get()
        dob = self.dob.get()
        gender = self.gender.get()
        baptised = self.baptised.get()
        residence = self.residence.get()
        department = self.department.get()
        tel = self.tel.get()
        nationality = self.nationality.get()
        email = self.email.get()
        image_data = self.image_path_var.get()
        if(name !=""):
            member_id = membership_backend.insert_member(memID, name, gender, dob, baptised, residence, department, tel, nationality, email)
            messagebox.showinfo("Success","Data is Successfully Stored",parent = self.addwin)
            with open(image_data, "rb") as f:
                image_data = f.read()

            membership_backend.insert_image_data(member_id, image_data)
            messagebox.showinfo("Success","Data is Successfully Stored",parent = self.addwin)
        else:
            messagebox.showerror("input error","Please all fields are required to be filled",parent = self.addwin)

        self.Reset()


    def Reset(self):
        self.gen_id()
        self.name.set("")
        self.dob.set("")
        self.gender.set("")
        self.baptised.set("")
        self.residence.set("")
        self.department.set("")
        self.tel.set("")
        self.nationality.set("")
        self.email.set("")
        self.image_path_var.set("")

    def count_male(self):
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM members WHERE gender='Male'")
        result = cur.fetchone()[0]
        conn.close()
        # Update the male count
        self.men.set(f"{result}")

    def count_female(self):
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM members WHERE gender='Female'")
        result = cur.fetchone()[0]
        conn.close()
        # Update the female count
        self.female.set(f"{result}")

    def count_Baptised(self):
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM members WHERE baptised='Yes'")
        result = cur.fetchone()[0]
        conn.close()
        # Update the female count
        self.bap.set(f"{result}")
    def count_all(self):
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM members")
        result = cur.fetchone()[0]
        self.allmem.set(f"{result}")


    def open_capture_window(self):
        capture_window = Toplevel(self.window)
        capture_window.config(bg = "#008080")
        capture_window.title("Photo Capture Window")
        capture_window.geometry("500x600+450+50")
        capture_window.iconbitmap('img/icon2.ico')
       
        
        self.photo_label = Label(capture_window)
        self.photo_label.pack()

        self.cap = cv2.VideoCapture(0)
        if not self.cap.isOpened():
            print("Error: Could not open camera.")
            return
        self.face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
        self.update_camera_feed()
        reset_button = Button(capture_window, text="Reset Camera",bg = "red",fg = 'white',font = ('Montserrat',14),width = 15 ,command=lambda: self.reset_camera_feed())
        reset_button.place(relx = .1,rely = .86)
        capture_button = Button(capture_window, text="Capture Photo",bg = "green",fg = 'white',font = ('Montserrat',14),width = 15 ,command=lambda: self.take_and_display_photo(capture_window))
        capture_button.place(relx =.5,rely =.86)

    def reset_camera_feed(self):
        self.cap.release()
        self.cap = cv2.VideoCapture(0)
        if not self.cap.isOpened():
            messagebox.showerror("Error","Could not open camera.")
            return
        self.update_camera_feed()

    def closeMain(self):
        self.window.destroy()
        self.exe_path = r'login.exe'
        subprocess.run(self.exe_path, check=True)

        

    def update_camera_feed(self):
        ret, frame = self.cap.read()

        if ret:
            faces = self.face_cascade.detectMultiScale(frame, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))
            for (x, y, w, h) in faces:
                cv2.rectangle(frame, (x, y), (x+w, y+h), (255, 0, 0), 2)
            rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            photo = ImageTk.PhotoImage(image=Image.fromarray(rgb_frame))
            self.photo_label.configure(image=photo)
            self.photo_label.image = photo

        self.window.after(5, self.update_camera_feed)

    def take_and_display_photo(self, window):
        self.cap.release()
        self.member_photos_folder = 'MemberPhotos'
        ret, frame = cv2.VideoCapture(0).read()
        image_filename = os.path.join(self.member_photos_folder, f"captured_photo_{len(os.listdir(self.member_photos_folder)) + 1}.png")
        os.makedirs(self.member_photos_folder, exist_ok=True)  # Create folder if not exists
        cv2.imwrite(image_filename, frame)
        # self.display_image(image_filename, self.photo)

    
    def bulkSMS(self):
        api_url = "https://sms.arkesel.com/sms/api?action=send-sms"
        api_key = "OjBxSFBoQ1NrUFJ6Q0MwR0s="

        phone_numbers = self.msgline.get().split(',')
        sender_id = self.id.get()
        message = self.body_entry.get("1.0", "end").strip()

        if not (phone_numbers and sender_id and message):
            messagebox.showwarning("Warning", "Please fill out all fields.", parent=self.msgwin)
            return

        success_count = 0  # Counter for successful sent messages

        for phone_number in phone_numbers:
            try:
                response = requests.get(f"{api_url}&api_key={api_key}&to={phone_number}&from={sender_id}&sms={message}")

                if response.status_code == 200:
                    success_count += 1
                else:
                    messagebox.showerror("Error", f"Failed to send SMS to {phone_number}. Status Code: {response.status_code}", parent=self.msgwin)
            except requests.exceptions.RequestException as e:
                messagebox.showerror("Error", f"Error sending SMS to {phone_number}: {e}", parent=self.msgwin)

        messagebox.showinfo("Success", f"{success_count} out of {len(phone_numbers)} SMS sent successfully!", parent=self.msgwin)
    def load_contacts_from_excel(self):
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")],parent=self.messagingFrame)
            if file_path:
                try:
                    df = pd.read_excel(file_path, dtype={'Phone': str})
                    contacts = df['Phone'].tolist()  
                    self.msgline.set(','.join(contacts))
                    messagebox.showinfo('Success', 'Contacts loaded successfully!',parent = self.messagingFrame)
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to load contacts: {e}',parent=self.messagingFrame)

    def search_attendace(self):
        search_string_data = self.search_string_attend.get() 
        if search_string_data:
            self.treeview.delete(*self.treeview.get_children())
            conn = sqlite3.connect('attendance_database.db')
            query = f"SELECT * FROM attendance WHERE date LIKE ? OR program LIKE ? OR male LIKE ? OR female LIKE ?"
            cursor = conn.execute(query, ('%' + search_string_data + '%', '%' + search_string_data + '%', '%' + search_string_data + '%', '%' + search_string_data + '%'))
            fetch = cursor.fetchall()
            count = 0
            for data in fetch:
                if count % 2 == 0:
                    self.treeview.insert('', 'end', values=data, tags=('evenrow'))
                else:
                    self.treeview.insert('', 'end', values=data, tags=('oddrow'))
                count += 1
            cursor.close()
            conn.close()
    
    def sheet_Tithe(self):
        cols = ["S/N", "mem_id",  "Full Name","Date Of Payment","Amount Paid", "Contact"]
        path = 'read.csv'
        excel_name = 'all_student_data{}.xlsx'.format(random.randint(1,1000))
        lst = []
        with open(path, "w", newline='') as myfile:
            csvwriter = csv.writer(myfile,delimiter=',')
            for row_id in self.treeview.get_children():
                row = self.treeview.item(row_id, 'values')
                lst.append(row)
            lst = list(map(list, lst))
            lst.insert(0, cols)
            for row in lst:
                csvwriter.writerow(row)

        writer = pd.ExcelWriter(excel_name)
        df = pd.read_csv(path)
        df.to_excel(writer, 'sheetname{}'.format(random.randint(1,10)))
        writer._save()
        messagebox.showinfo("export message","Data Exported Successfully")

    def delete_tithe(self):
        if self.treeview.selection():
            result = messagebox.askquestion('Python - Delete Data Row In SQLite',
                                            'Are you sure you want to delete this record?', icon="warning",parent = self.givewin)
            if result == 'yes':
                curItem = self.treeview.focus()
                contents = (self.treeview.item(curItem))
                selecteditem = contents['values']
                self.treeview.delete(curItem)
                giving_backend.delete_Tithe(selecteditem[0])

                self.DisplayData_Tithe()

            else:
                self.DisplayData_Tithe()

    def generate_pdf_tithe(self):
        doc = SimpleDocTemplate("treeview.pdf",)
        data = []
        columns = []
        for col in self.treeview["columns"]:
            column_heading = self.treeview.heading(col)["text"]
            columns.append(column_heading)
        data.append(columns)
        for item in self.treeview.get_children():
            values = self.treeview.item(item)["values"]
            data.append(values)
        table_data = Table(data)
        styles = getSampleStyleSheet()
        heading_style = styles["Heading1"]
        heading_style.alignment = 1 
        heading = Paragraph("<b>DATASHEET</b>", heading_style)

        elements = [heading, Spacer(1, 20), table_data]
        doc.build(elements)
        subprocess.run(["start", "treeview.pdf"], shell=True)    
    
    def search_data_Tithe(self):
        search_string_data = self.search_string_tithe.get() 
        if search_string_data:
            self.treeview.delete(*self.treeview.get_children())
            conn = sqlite3.connect('GiveDB.db')
            query = f"SELECT * FROM tithe WHERE memID LIKE ? OR name LIKE ? OR date LIKE ? OR contact LIKE ?"
            cursor = conn.execute(query, ('%' + search_string_data + '%', '%' + search_string_data + '%', '%' + search_string_data + '%', '%' + search_string_data + '%'))
            fetch = cursor.fetchall()
            count = 0
            for data in fetch:
                if count % 2 == 0:
                    self.treeview.insert('', 'end', values=data, tags=('evenrow'))
                else:
                    self.treeview.insert('', 'end', values=data, tags=('oddrow'))
                count += 1
            cursor.close()
            conn.close()

    def DisplayData_Tithe(self):
            self.treeview.delete(*self.treeview.get_children())
            conn = sqlite3.connect('GiveDB.db')
            cursor = conn.execute("SELECT * FROM tithe")
            fetch = cursor.fetchall()
            count = 0
            for data in fetch:
                if count % 2 == 0:
                    self.treeview.insert('', 'end', values=(data), tags=('evenrow'))
                else:
                    self.treeview.insert('', 'end', values=(data), tags=('oddrow'))
                count += 1

            cursor.close()
            conn.close()

    def visitors(self):
        self.main_visitors = Frame(self.attendwin,width=1095,height = 600,bg = 'white')
        self.main_visitors.place(relx =0,rely = .11)

        self.close_command = Button(self.main_visitors,text= 'close',bg = 'red',fg = 'white',font= ('Arial',14),command = lambda: self.main_visitors.destroy())
        self.close_command.place(relx = 0.9,rely = 0)
        

        self.visi_register = Frame(self.main_visitors,width = 700,height = 200,bg = 'white',bd = 2,relief=SUNKEN)
        self.visi_register.place(relx = 0.15,rely = 0.1)

        self.name = Label(self.visi_register,text = "Name",font = ("Montserrat",14),bg ='white',)
        self.name.place(relx = 0.01,rely = 0.0)
        self.nameEntry = Entry(self.visi_register,font=("Montserrat",14),width = 22,bd = 2)
        self.nameEntry.place(relx = 0.15,rely = 0)

        self.gender = Label(self.visi_register,text = "Gender",font = ("Montserrat",14),bg ='white',)
        self.gender.place(relx = 0.6,rely = 0.0)
        self.genderEntry = ttk.Combobox(self.visi_register,font=("Montserrat",14),width = 12,values=("Male","Female"))
        self.genderEntry.place(relx = 0.75,rely = 0)

        self.address = Label(self.visi_register,text = "Address",font = ("Montserrat",14),bg ='white')
        self.address.place(relx = 0.01,rely = 0.2)
        self.addressEntry = Entry(self.visi_register,font=("Montserrat",14),width = 22,bd = 2)
        self.addressEntry.place(relx = 0.15,rely = .2)

        self.contact = Label(self.visi_register,text = "Contact",font = ("Montserrat",14),bg ='white',)
        self.contact.place(relx = 0.6,rely = 0.2)
        self.contactEntry = Entry(self.visi_register,font=("Montserrat",14),width = 14,bd = 2)
        self.contactEntry.place(relx = 0.75,rely = .2)

        self.invitedby = Label(self.visi_register,text = "Invited By",font = ("Montserrat",14),bg ='white',)
        self.invitedby.place(relx = 0.01,rely = 0.4)
        self.invitedbyEntry = Entry(self.visi_register,font=("Montserrat",14),width = 22,bd = 2)
        self.invitedbyEntry.place(relx = 0.15,rely = .4)

        self.dateVisited = Label(self.visi_register,text = "Date",font = ("Montserrat",14),bg ='white',)
        self.dateVisited.place(relx = 0.6,rely = 0.4)
        self.dateVisitedEntry = DateEntry(self.visi_register,font=("Montserrat",14),width = 12,date_pattern = "dd-mm-yyyy")
        self.dateVisitedEntry.place(relx = 0.75,rely = .4)

        self.purpose = Label(self.visi_register,text = "Purpose",font = ("Montserrat",14),bg ='white',)
        self.purpose.place(relx = 0.01,rely = 0.6)
        self.purposeEntry = Entry(self.visi_register,font=("Montserrat",14),width = 35,bd = 2)
        self.purposeEntry.place(relx = 0.15,rely = .6)

        # Button

        self.submit = Button(self.main_visitors,width = 30,font = ("Montserrat",18,'bold'),text = "Submit",bg = '#008080',fg = 'white')
        self.submit.place(relx = .3,rely =.5)




win = Tk()

obj = CManagement(win)
win.iconbitmap("img/icon2.ico")
win.mainloop()