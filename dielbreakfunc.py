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
from matplotlib.ticker import FormatStrFormatter
import matplotlib.pyplot as plt
import numpy as np



class WinDielBreak:
    def __init__(self, root,title,fname,const):
        self.root = root
        self.root.geometry("+150+50")
        self.fname =fname
        self.const =const
        self.title=title
        self.root.wm_title(self.title)
        self.sMin1=StringVar()
        self.sMin1.set("-999")
        self.sMax1=StringVar()
        self.sMax1.set("999")
        
        frame = Frame(self.root, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        frameD1 = Frame(self.root)
        frameD1.pack(fill=X, expand=True)
        frameD2 = Frame(self.root)
        frameD2.pack(fill=X, expand=True)
        
        self.frame1 = Frame(frame)
        self.frame1.pack(fill=BOTH)
        self.fig = Figure(figsize=(7, 6), dpi=100)

        dd=DielBreak(self.fname,[float(self.sMin1.get()),float(self.sMax1.get())])
        b=fname.split("DielectricBreak")[1].split(".t")[0]
        dirD=os.path.dirname(os.path.abspath(fname))
        self.img=dirD+"/DielectricBreak_flute2_img"+b+".pdf"
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[2])
        self.result.append(dd[3])
        self.result.append(dd[4])
        
        self.ax=self.fig.add_subplot(111)
        self.ax.plot(dd[0],dd[1],".")
        self.ax.axhline(y=1,c="orange",ls='dashed')
        self.ax.axvline(x=dd[2],c="green")
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.xaxis.get_offset_text().set_position((1,0.1))
        self.ax.xaxis.get_offset_text().set_rotation(-90)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("Current [nAF]",loc='top')
        self.ax.text(0.5,0.2, ("Result:\n%.2f$\pm$%.2f [V]" % (dd[2],dd[3])), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.frame1)  # A tk.DrawingArea.
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.frame1)
        self.toolbar.update()
        self.canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

        self.lem1=Label(frameD1,text="Min1:", anchor=W,justify=LEFT, width=6)
        self.lem1.pack(side=LEFT, padx=5, pady=5)
        self.em1=Entry(frameD1, textvariable=self.sMin1, width=10)
        self.em1.pack(side=LEFT)#, padx=5, pady=5)

        self.leM1=Label(frameD1,text="Max1:", anchor=W,justify=LEFT, width=6)
        self.leM1.pack(side=LEFT, padx=5, pady=5)
        self.eM1=Entry(frameD1, textvariable=self.sMax1, width=10)
        self.eM1.pack(side=LEFT)#, padx=5, pady=5)
                
        self.b3=Button(self.root,text='Next',command=self.closeExt)
        self.b3.pack(side=RIGHT, padx=5, pady=5)
        self.b2=Button(self.root,text='Update',command= self.UpdateDielBreak)
        self.b2.pack(side=RIGHT)

        self.root.wait_window(self.root)
        
    def closeExt(self):
        self.fig.savefig(self.img)
        self.root.destroy()

    def UpdateDielBreak(self):
        dd=DielBreak(self.fname,[float(self.sMin1.get()),float(self.sMax1.get())])
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[2])
        self.result.append(dd[3])
        self.result.append(dd[4])

        self.ax.clear()

        self.ax.plot(dd[0],dd[1],".")
        self.ax.axhline(y=1,c="orange",ls='dashed')
        self.ax.axvline(x=dd[2],c="green")
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.ticklabel_format(axis='x',style='sci',scilimits=(0,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("Current [nA]",loc='top')
        self.ax.text(0.5,0.2, ("Result:\n%.2f$\pm$%.2f [pF]" % (dd[2],dd[3])), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        self.canvas.draw()
        self.toolbar.update()







def DielBreak(fname,limits):
    f=open(fname)
    d=f.read().splitlines()
    d.pop(0)
    x=[]
    y=[]
    for k in d:
        item=k.split("\t")
        x.append(float(item[0]))
        y.append(float(item[1])*1e9)

    xx=np.array(x)
    yy=np.array(y)
    
    min1=limits[0]
    max1=limits[1]


    try:
        idmin1=int(np.where(xx==min1)[0])
    except:
        idmin1=0

    try:
        idmax1=int(np.where(xx==max1)[0])
    except:        
        idmax1=xx.size


    resId=-1        
    for idx in range(idmin1,idmax1):
        if y[idx]> 1:
            resId=idx
            break
    if resId!=-1:
        val1=x[resId]
        err1=val1*0.01
        flag=""
    else:
        val1=200
        err1=0
        flag=">"

    return xx,yy,val1,err1,flag

      
