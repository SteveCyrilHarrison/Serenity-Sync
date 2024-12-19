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


class Serenity:
    def __init__(self,window):
        self.window = window
        self.window.geometry("1360x760+0+0")
        self.window.title("SerentySync")
        self.window.resizable(0,0)
        # self.window.after(0,self.update_count)
        self.user = StringVar()
        self.user.set("USER-ADMIN")
        self.allmem =StringVar()
        self.men = StringVar()
        self.female = StringVar()
        self.bap = StringVar()

        self.DashBoard()

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



win = Tk()

obj = Serenity(win)
win.iconbitmap("img/icon2.ico")
win.mainloop()
