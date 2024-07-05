
from flask import Flask,render_template,redirect,request
import pywhatkit
import openpyxl
from tkinter import *
from tkinter import filedialog

times ="V:\python\python_whatsapp_automation\images\LOGO_COCK_1-01-01.png" 
load = openpyxl.load_workbook("test.xlsx")
sheet = load.active


window = Tk()


window.title('File Explorer')

window.geometry("500x500")

window.config(background = "white")


def browseFiles():
    global  filename
    filename = filedialog.askopenfilename(initialdir = "/",
										title = "Select a File",
										filetypes = (("image",
														".png"),
													("all files",
														".")))
	

    label_file_explorer.configure(text="File Opened: "+filename)
   
def image_send():
         i=1
         while i<3:
            number = sheet[f'A{i}'].value
            print(f"+91{number}")
            # pywhatkit.sendwhatmsg_instantly(f"+91{number}", "Hi",5)
            pywhatkit.sendwhats_image(f"+91{number}", f"{filename}")
            i+=1

	
																								

Label(window, text='First Name').grid(row=8,column=1)
Label(window, text='Last Name').grid(row=9,column=1)
e1 = Entry(window)
e2 = Entry(window)
e1.grid(row=8, column=1)
e2.grid(row=9, column=1)

label_file_explorer = Label(window, 
							text = "File Explorer using Tkinter",
							width = 100, height = 4, 
							fg = "blue")

	
button_explore = Button(window, 
						text = "Browse Files",
						command = browseFiles) 
	
button_send = Button(window, 
						text = "send file",
						command =image_send) 
button_exit = Button(window, 
					text = "Exit",
					command = exit) 


label_file_explorer.grid(column = 1, row = 1)

button_explore.grid(column = 1, row = 2)
button_send.grid(column=1,row=5)
button_exit.grid(column = 1,row = 3)


window.mainloop()