import os
import re
import time
import pdfkit
from tkinter import *
import tkinter as tk
from fpdf import FPDF
import win32com.client
from tkinter import filedialog, ttk
from tkinter.ttk import Progressbar
from tkinterdnd2 import DND_FILES, TkinterDnD
from docx2pdf import  convert

global path
global ChoixFct
global progressPCT
filesListToConvert = []  #a global list to store path of files
ChoixFct =' Word to Pdf' #default value
progressPCT = 0

# ppt to pdf function
def PPT_to_PDF(infile_path, outfile_path):
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    pdf = powerpoint.Presentations.Open(infile_path, WithWindow=False)
    pdf.SaveAs(outfile_path, 32)
    pdf.Close()
    powerpoint.Quit()

# function to change converter option
def changeConverterFunction(event):
    global ChoixFct
    ChoixFct = event.widget.get()

# function to put in lambda function (anonymous)
def lambdaFunct(e):
    links = re.findall(r'(\/.*?\.[\w:]+)', e.data)
    match ChoixFct:
        case ' Word to Pdf':
            for link in links:
                file_extension = os.path.splitext(link)
                if (file_extension[1] == '.docx' or file_extension[1] == '.docm' or file_extension[1] == '.doc'):
                    lb.insert(tk.END, link)
                    filesListToConvert.append(link)
        case ' PPT to Pdf':
            for link in links:
                file_extension = os.path.splitext(link)
                if (file_extension[1] == '.pptx' or file_extension[1] == '.ppt' or file_extension[1] == '.odp'):
                    lb.insert(tk.END, link)
                    filesListToConvert.append(link)
        # case ' EXC to Pdf':
        #     for link in links:
        #         file_extension = os.path.splitext(link)
        #         if (file_extension[1] == '.xls'):
        #             lb.insert(tk.END, link)
        #             filesListToConvert.append(link)
        case ' IMG to Pdf':
            for link in links:
                file_extension = os.path.splitext(link)
                if (file_extension[1] == '.jpg' or file_extension[1] == '.png' or file_extension[1] == '.jpeg'):
                    lb.insert(tk.END, link)
                    filesListToConvert.append(link)
        case ' TXT to Pdf':
            for link in links:
                file_extension = os.path.splitext(link)
                if (file_extension[1] == '.txt'):
                    lb.insert(tk.END, link)
                    filesListToConvert.append(link)

# function for getting saving path (if you dont chose it is the same place where your file is located)
# function of converting the list of files included
def pathASK():
    path = filedialog.askdirectory()
    match ChoixFct:
        case ' Word to Pdf':
            for name in filesListToConvert:
                convert(name,path)
        case ' PPT to Pdf':
            for name in filesListToConvert:
                # print(name)
                cachePathName = path
                basename = os.path.basename(name)
                pdfFileName = os.path.splitext(basename)[0]
                path = path + "/" + pdfFileName + ".pdf"
                PPT_to_PDF(name,path.replace("/", "\\\\"))
                path = cachePathName
        # case ' EXC to Pdf':
        #     for name in filesListToConvert:
        #         convert(name,path)
        case ' IMG to Pdf':
            for name in filesListToConvert:
                pdf = FPDF()
                cachePathName = path
                basename = os.path.basename(name)
                pdfFileName = os.path.splitext(basename)[0]
                path = path + "/" + pdfFileName + ".pdf"
                pdf.add_page()
                pdf.image(name, 0,0,210,297)
                pdf.output(path, 'F')
                path = cachePathName
        case ' TXT to Pdf':
            for name in filesListToConvert:
                cachePathName = path
                with open(name) as file:
                    basename = os.path.basename(name)
                    pdfFileName = os.path.splitext(basename)[0]
                    pdfFileName = pdfFileName + ".pdf"
                    path = path + "/" +pdfFileName
                    print(path)
                    with open (path, "w") as output:
                        file = file.read()
                        file = file.replace("\n", "<br>")
                        output.write(file)
                path = cachePathName

# hover the buttons effect
def hoverActive(boton, color1, color2, color3):
	boton.configure(bg=color1)
	def fuera(e):
		boton.configure(bg=color1)
	def dentro(e):
		boton.configure(bg=color2)
	def activo(e):
		boton.configure(activebackground=color3)
	boton.bind("<Enter>", dentro)
	boton.bind("<Leave>", fuera)
	boton.bind("<ButtonPress-1>", activo)

# progressbar function
def startProcessing():
    global progressPCT
    for i in range(1,40):
        progress['value'] = i
        progressPCT = i
        root.update_idletasks()
        value_label['text'] = update_progress_label()
        time.sleep(0.004)
    pathASK()
    for i in range(40,101):
        progress['value'] = i
        progressPCT = i
        root.update_idletasks()
        value_label['text'] = update_progress_label()
        time.sleep(0.002)
    finishWindow()

# percentage of progressbar
def update_progress_label():
    global progressPCT
    return f"Current Progress: {progressPCT}%"

def finishWindow():
    secondWindow_width = 190
    secondWindow_height = 100
    secondWindow = tk.Toplevel()
    secondWindow.title("Files Converter")
    secondWindow.iconbitmap('./logo/logo.ico')
    secondWindow.resizable(False,False)
    secondWindowx = (screen_width / 2) - (secondWindow_width / 2)
    secondWindowy = (screen_height / 2 ) - (secondWindow_height / 2)
    secondWindow.geometry(f'{secondWindow_width}x{secondWindow_height}+{int(secondWindowx)}+{int(secondWindowy)}')
    secondWindow.grab_set()
    labelSS = tk.Label(secondWindow, text="Files Converted Successfully")
    labelSS.pack(pady=10)
    def close_second_window():
        secondWindow.grab_release()
        secondWindow.destroy()
    buttonSS = tk.Button(secondWindow, text="OK", command=close_second_window)
    buttonSS.pack(pady=10)

# define of gui of app
root = TkinterDnD.Tk()
root.title('Files Converter')
root.iconbitmap('./logo/logo.ico')
# Designate Height and Width of our app
app_width = 750
app_height = 350
# The Height and Width of our pc screen
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width / 2) - (app_width / 2)
y = (screen_height / 2 ) - (app_height / 2)
root.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')
label_1 = Label(root,width="550",height="285",bg="#EE4E34")#to colorate the space of application
label_1.place(x=0,y=0)
root.resizable(False,False)
ChoixFonction = ttk.Combobox(root,width=35, cursor="hand2", state="readonly", foreground= "#000")
Btnpath = Button(root,fg= "#000",text="Convert Here", cursor="hand2",borderwidth=0,width=35, command=startProcessing)
hoverActive(Btnpath, "#ffffff", "#FCEDDA", "#ffffff")
ChoixFonction['values'] = (' Word to Pdf',' PPT to Pdf',' TXT to Pdf',' IMG to Pdf') 
ChoixFonction.pack(ipadx=5, ipady=5, padx=6, pady=4)
Btnpath.pack( ipadx=5, ipady=5, padx=6, pady=4)
ChoixFonction.current(0) #default value
ChoixFonction.bind("<<ComboboxSelected>>",changeConverterFunction)

lb = tk.Listbox(root, width=120, bd=1,height=13,selectbackground= "#EE4E34", cursor="hand2" , bg="#FCEDDA")
# register the listbox as a drop target
lb.drop_target_register(DND_FILES)
lb.dnd_bind('<<Drop>>', lambda e: lambdaFunct(e))
lb.pack()

progress = Progressbar(root, length=500, mode='determinate')
progress.pack(pady=5,padx=5)
value_label = ttk.Label(root, text=update_progress_label())
value_label.pack(pady=1,padx=1)

root.mainloop()