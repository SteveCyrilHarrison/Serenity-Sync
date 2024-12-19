from tkinter import *
from PIL import ImageTk,Image
from tkinter import messagebox
from tkinter.ttk import Progressbar
import sys
import subprocess

class LoadWin:
    def __init__(self,win):
        self.win = win
        self.win.geometry("800x500+300+100")
        self.win.overrideredirect(True)


        # variables
        self.userID = StringVar()
        self.passcode = StringVar()
        
        
        self.LoadingPage()

    def LoadingPage(self):
        self.loadingFrame = Frame(self.win, width = 800,height = 500,bg = 'white')
        self.loadingFrame.place(relx = 0,rely =0)

        self.image_lbl = Label(self.loadingFrame,bg = 'white')
        self.logo = Image.open('img/logo.jpg')
        self.resized = self.logo.resize((450,400))
        self.real_image = ImageTk.PhotoImage(self.resized)
        self.image_lbl.config(image =self.real_image)
        self.image_lbl.place(relx = 0.23,rely= 0)
        # self.church_label = Label(self.loadingFrame,text = "ICGC ROMAN DOWN ASSEMBLY",font = ('Montserrat',15),bg = "white")
        # self.church_label.place(relx= .3,rely = .54)
        self.progress_label = Label(self.loadingFrame,text = "loading Data...",bg="white",font = ('Montserrat',14,"bold"),fg ="#008080")
        self.progress_label.place(relx = 0.35,rely = 0.75)
        self.progress_bar = Progressbar(self.loadingFrame,length = 670,orient =HORIZONTAL,mode = 'determinate')
        self.progress_bar.place(relx = 0.1,rely =0.88)

        def exit_window():
            sys.exit(self.loadingFrame.destroy())

#=================================================progressbar function=====================================================
        global i
        i=0
        def load():
            global i
            if i <=10:
                txt="Fetching System Data..."+(str(10*i)+'%')
                self.progress_label.config(text=txt)
                self.progress_label.after(1000,load)
                self.progress_bar['value'] = 10*i
                i += 1

                if self.progress_bar['value']==100:
                    self.win.destroy()
                    self.exe_path = r'login.exe'
                    subprocess.run(self.exe_path, check=True)



        load()

root = Tk()
obj = LoadWin(root)
root.mainloop()