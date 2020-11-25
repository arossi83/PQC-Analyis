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



class WinDiodeIV:
    def __init__(self, root,title,fname,const,save=True):
        self.root = root
        self.root.geometry("+150+50")
        self.fname =fname
        self.title=title
        self.const=const
        self.save=save
        self.root.wm_title(self.title)
        self.sMin1=StringVar()
        self.sMin1.set("600")
        
        frame = Frame(self.root, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        frameD1 = Frame(self.root)
        frameD1.pack(fill=BOTH, expand=True)
        
        self.frame1 = Frame(frame)
        self.frame1.pack(fill=X)
        self.fig = Figure(figsize=(7, 6), dpi=100)

        dd=diodeiv(self.fname,float(self.sMin1.get()))
        self.sMin1.set(dd[5])
        
        if "DiodeIV" in fname:
            b=fname.split("DiodeIV")[1].split(".t")[0]
        else:
            b=fname.split("Diode_IV")[1].split(".t")[0]            
        dirD=os.path.dirname(os.path.abspath(fname))
        self.img=dirD+"/DiodeIV_flute3_img"+b+".pdf"
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[2])
        self.result.append(dd[3])
        self.result.append(dd[4])

        self.ax=self.fig.add_subplot(111)
        self.ax.plot(dd[0],dd[1],".")
        self.ax.axhline(y=dd[2],c="orange")
        self.ax.axvline(x=float(dd[4].split("@")[1].split("V")[0]),c="green",ls='dashed')
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("Current [A]",loc='top')
        self.ax.set_ylim(dd[1][int(np.where(dd[0]==100)[0])]*0.9,dd[1][-1]*1.2)
        self.ax.text(0.4,0.5, ("Result:\n%.2f$\pm$%.2f [pA]" % (dd[2]*1e12,dd[3]*1e12)), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.frame1)  # A tk.DrawingArea.
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.frame1)
        self.toolbar.update()
        self.canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

        self.lem1=Label(frameD1,text="Value:", justify=RIGHT, width=4)
        self.lem1.grid(column=0,row=0)
        self.em1=Entry(frameD1, textvariable=self.sMin1,justify=LEFT,width=7)
        self.em1.grid(column=1,row=0)


                
        self.b3=Button(self.root,text='Next',command=self.closeExt)
        self.b3.pack(side=RIGHT, padx=5, pady=5)
        self.b2=Button(self.root,text='Update',command= self.UpdateDiodeIV)
        self.b2.pack(side=RIGHT)

        self.root.wait_window(self.root)
        
    def closeExt(self):
        if self.save:
            self.fig.savefig(self.img)
        self.root.destroy()

    def UpdateDiodeIV(self):
        dd=diodeiv(self.fname,float(self.sMin1.get()))
        self.sMin1.set(dd[5])
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[2])
        self.result.append(dd[3])
        self.result.append(dd[4])

        self.ax.clear()

        self.ax.plot(dd[0],dd[1],".")
        self.ax.axhline(y=dd[2],c="orange")
        self.ax.axvline(x=float(dd[4].split("@")[1].split("V")[0]),c="green",ls='dashed')
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("Current [A]",loc='top')
        self.ax.set_ylim(dd[1][int(np.where(dd[0]==100)[0])]*0.9,dd[1][-1]*1.2)
        self.ax.text(0.4,0.5, ("Result:\n%.2f$\pm$%.2f [pA]" % (dd[2]*1e12,dd[3]*1e12)), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        self.canvas.draw()
        self.toolbar.update()







def diodeiv(fname,limits):
    f=open(fname)
    d=f.read().splitlines()
    d.pop(0)
    d.pop(0)
    x=[]
    y=[]
    for k in d:
        item=k.split("\t")
        x.append(-1*float(item[0]))
        y.append(float(item[1]))

    xx=np.array(x)
    yy=np.array(y)

    flag=""
    try:
        val1=yy[int(np.where(xx==limits)[0])]
        flag=(("@%dV" % (limits)))
    except:
        val1=yy[int(np.where(xx==400)[0])]
        flag="@400V"
        limits=400

    err1=val1*0.01
        
    return xx,yy,val1,err1,flag,limits

      
