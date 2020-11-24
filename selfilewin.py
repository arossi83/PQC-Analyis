import tkinter
from tkinter import *
from tkinter import filedialog,StringVar
from tkinter.ttk import Frame, Button, Style
import sys
import os
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
# Implement the default Matplotlib key bindings.
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import numpy as np



class WinSel:
    def __init__(self, root,names,indices):
        self.root = root
        self.root.geometry("+200+200")

        self.val=IntVar()
        self.val.set(0)


        frame = Frame(self.root, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        for i,n in enumerate(indices):
            txt=names[n].split("/")[-1]
            print(txt)
            self.rad=Radiobutton(frame, text=txt, variable=self.val, value=i,width=80, relief=GROOVE).pack(anchor=W)


        
        self.b3=Button(self.root,text='Next',command=self.closeExt)
        self.b3.pack(side=RIGHT, padx=5, pady=5)

        self.root.wait_window(self.root)

        
    def closeExt(self):
        print("Selection : %d" % (self.val.get()))
        self.root.destroy()

