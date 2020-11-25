import tkinter
from tkinter import *
from tkinter import filedialog,StringVar
from tkinter.ttk import Frame, Button, Style
import xlsxwriter
import sys
import os
import csv
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
# Implement the default Matplotlib key bindings.
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import numpy as np
from diodeCVfunc import WinDiodeCV
from selfilewin import WinSel


mainDir="/Users/alessandro/Documents/CMS/PQC_Analysis"

def SelFileWindow(Win_class,names,indices):
    win2 = Toplevel(root)
    q=Win_class(win2,names,indices)
    return q.val.get()
    
def new_window(Win_class,title,fname,const,save):
    win2 = Toplevel(root)
    q=Win_class(win2,title,fname,const,save)
    return q.result
    
class mainWindow(object):
    def __init__(self,master):
        self.master=master

        frame = Frame(master, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)

        frame1 = Frame(frame)
        frame1.pack(fill=X)
        frame2 = Frame(frame)
        frame2.pack(fill=X)
        frame3 = Frame(frame)
        frame3.pack(fill=X)

        self.lf=Label(frame1,textvariable=dname, anchor=W, justify=CENTER)
        self.lf.pack(side=LEFT, padx=5, pady=25, expand=True)        

        self.b1=Button(frame2,text="Select File",command=self.LoadFile)
        self.b1.pack(side=RIGHT, padx=5, expand=True)

        self.outChk=Checkbutton(frame3, text="Save Fig", anchor=W, var=saveFig)
        self.outChk.pack()
        
        self.b3=Button(master,text='Close',command=self.cleanup)
        self.b3.pack(side=RIGHT, padx=5, pady=5)
        self.b2=Button(master,text='Process',command=self.process)
        self.b2.pack(side=RIGHT)

    def process(self):
        Res=[]
        print("Analyze DiodeCV data")
        Res.append(new_window(WinDiodeCV,"Diode C/V Depletion Voltage",dname.get()),1,saveFig.get())
        print("DiodeCV done")


        
    def cleanup(self):
        self.master.destroy()
        
    def LoadFile(self):   
        dname.set(filedialog.askopenfilename(initialdir=mainDir,filetypes = (("text files","*.txt"),("all files","*.*"))))

        
if __name__ == "__main__":
    root=Tk()
    root.title("Diode C/V Analysis")

    dname=StringVar()
    dname.set("Please select a File")
    saveFig=StringVar()
    saveFig.set(True)
    outfname=StringVar()
    outfname.set("")

    m=mainWindow(root)
    root.after(100, lambda: root.focus_force())
    root.mainloop()

    
