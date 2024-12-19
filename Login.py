from tkinter import *
from PIL import ImageTk,Image
from tkinter import messagebox
import subprocess

class Login_Window:
    def __init__(self,win):
        self.win = win
        self.win.geometry("800x500+300+100")
        self.win.overrideredirect(True)


        # variables
        self.userID = StringVar()
        self.passcode = StringVar()
        
        
        self.load_contents()


    def load_contents(self):
        self.sign()

    def sign(self):
        
        self.content = Frame(self.win,width = 360, height = 500,bg= "white")
        self.content.place(relx = .55,rely =0)

        self.image_lbl = Label(self.content,bg = 'white')
        self.logo = Image.open('img/images.png')
        self.resized = self.logo.resize((160,160))
        self.real_image = ImageTk.PhotoImage(self.resized)
        self.image_lbl.config(image =self.real_image)
        self.image_lbl.place(relx = 0.23,rely= 0.05)
        # contents
        self.userid = Label(self.content,text = "Membership ID", font = ('Montserrat',14),bg = 'white')
        self.userid.place(relx = .02,rely =.4)
        self.userid_field = Entry(self.content,font = ('Montserrat',14),width = 25,bd = 2,textvariable=self.userID)
        self.userid_field.place(relx=.02, rely = .46)

        self.userpass = Label(self.content,text = "Passcode", font = ('Montserrat',14),bg = 'white')
        self.userpass.place(relx = .02,rely =.55)
        self.userpass_field = Entry(self.content,font = ('Montserrat',14),width = 25,bd = 2,show = '*',textvariable=self.passcode)
        self.userpass_field.place(relx=.02, rely = .62)

        self.loginButton = Button(self.content,text='Login',font = ('Montserrat',14,'bold'),width = 25,bg= "#008080",fg = "white",command = self.Loginlogic)
        self.loginButton.place(relx = .02, rely = .74)

        self.loginButton = Button(self.content,text='cancel',font = ('Montserrat',12,'bold'),bg = "white",width = 25,bd= 0,fg = "red",command = self.cancel)
        self.loginButton.place(relx = .1, rely = .85)

        self.content2 = Frame(self.win,width = 440, height = 500,bg= "white")
        self.content2.place(relx =0,rely =0)

        self.printer = Label(self.content2)
        self.logop = Image.open('img/church.jpg')
        self.resizedp= self.logop.resize((430,600))
        self.real_imagep =ImageTk.PhotoImage(self.resizedp)
        self.printer.config(image =self.real_imagep)
        self.printer.place(relx = 0,rely= 0)

    def Loginlogic(self):
        if(self.userID.get())=="Admin001" and (self.passcode.get())=="admin001":
            messagebox.showinfo("LoggedIn","Login Successful")
            self.win.destroy()
            self.exe_path = r'SerenitySync.exe'
            subprocess.run(self.exe_path, check=True)
            
        
        else:
            messagebox.showerror("Error", "Invalid Login Credentials")
    def cancel(self):
        self.win.destroy()
            
        


main_win = Tk()
obj = Login_Window(main_win)
main_win.mainloop()