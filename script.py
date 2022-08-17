from tkinter import *
from tkinter import filedialog
import re
import os
import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
from docx2pdf import  convert

global path
filesListToConvert = [] #a global list to store path of files

# def convertFunction():
#     for name in filesListToConvert:
#         # convert(r"C:\Users\Hp\OneDrive\Bureau\test.docx",r"C:\Users\Hp\OneDrive\Bureau\test.pdf")
#         print(name)

# function to put in lambda function (anonymous)
def lambdaFunct(e):
    links = re.findall(r'(\/.*?\.[\w:]+)', e.data)
    for link in links:
        file_extension = os.path.splitext(link)
        if (file_extension[1] == '.docx' or file_extension[1] == '.docm' or file_extension[1] == '.doc'):
            lb.insert(tk.END, link)
            filesListToConvert.append(link)

# function for getting saving path (if you dont chose it is the same place where your file is located)
# function of converting the list of files included
def pathASK():
    if 1==1:
        path = filedialog.askdirectory()
        for name in filesListToConvert:
            convert(name,path)
            # print(name)

# define of gui of app
root = TkinterDnD.Tk()
root.title('List Files Converter')
root.iconbitmap('logo.ico')
# Designate Height and Width of our app
app_width = 750
app_height = 300
# The Height and Width of our pc screen
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width / 2) - (app_width / 2)
y = (screen_height / 2 ) - (app_height / 2)
root.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')
label_1 = Label(root,width="550",height="285",bg="#EE4E34")#to colorate the space of application
label_1.place(x=0,y=0)
root.resizable(False,False)
Btnpath = Button(root,fg= "#000",text="Convert Here",borderwidth=0,width=35, command=pathASK)
Btnpath.pack( ipadx=5, ipady=5, padx=6, pady=4)


lb = tk.Listbox(root, width=120, bd=1,height=15,selectbackground= "#EE4E34", cursor="hand2" , bg="#FCEDDA")
# register the listbox as a drop target
lb.drop_target_register(DND_FILES)
lb.dnd_bind('<<Drop>>', lambda e: lambdaFunct(e))
lb.pack()


root.mainloop()