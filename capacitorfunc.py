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



class WinCap:
    def __init__(self, root,title,fname,const,save):
        self.root = root
        self.root.geometry("+150+50")
        self.fname =fname
        self.const =const
        self.title=title
        self.save=save
        self.root.wm_title(self.title)
        self.sMin1=StringVar()
        self.sMin1.set("-999")
        self.sMax1=StringVar()
        self.sMax1.set("999")
        
        frame = Frame(self.root, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        frameD1 = Frame(self.root)
        frameD1.pack(fill=X, expand=True)
        
        self.frame1 = Frame(frame)
        self.frame1.pack(fill=BOTH)
        self.fig = Figure(figsize=(7, 6), dpi=100)

        dd=Cap(self.fname,[float(self.sMin1.get()),float(self.sMax1.get())])
        b=fname.split("Capacitor")[1].split(".t")[0]
        dirD=os.path.dirname(os.path.abspath(fname))
        self.img=dirD+"/Capacitor_flute1_img"+b+".pdf"
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[2][0])
        self.result.append(dd[3][0])
        self.result.append("pF")
        
        line1=np.poly1d(([0,dd[2][0]]))
        x1 = np.linspace(dd[0][0],dd[0][-1], 100)

        
        self.ax=self.fig.add_subplot(111)
        self.ax.plot(dd[0],dd[1],".",x1,line1(x1))
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.xaxis.get_offset_text().set_position((1,0.1))
        self.ax.xaxis.get_offset_text().set_rotation(-90)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("Capacitance [pF]",loc='top')
        self.ax.text(0.5,0.2, ("Result:\n%.4f$\pm$%.4f [pF]" % (dd[2][0],dd[3][0])), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.frame1)  # A tk.DrawingArea.
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.frame1)
        self.toolbar.update()
        self.canvas.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)


        self.lem1=Label(frameD1,text="Min1:", font=('Helvetica',12),justify=RIGHT, width=4)
        self.lem1.grid(column=0,row=0)
        self.em1=Entry(frameD1, textvariable=self.sMin1,font=('Helvetica',12),justify=LEFT,width=7)
        self.em1.grid(column=1,row=0)
        self.leM1=Label(frameD1,text="Max1:", font=('Helvetica',12),justify=RIGHT, width=4)
        self.leM1.grid(column=2,row=0)
        self.eM1=Entry(frameD1, textvariable=self.sMax1, font=('Helvetica',12),justify=LEFT,width=7)
        self.eM1.grid(column=3,row=0)


        self.b3=Button(self.root,text='Next',command=self.closeExt)
        self.b3.pack(side=RIGHT, padx=5, pady=5)
        self.b2=Button(self.root,text='Update',command= self.UpdateCap)
        self.b2.pack(side=RIGHT)

        self.root.wait_window(self.root)
        
    def closeExt(self):
        if self.save:
            self.fig.savefig(self.img)
        self.root.destroy()

    def UpdateCap(self):
        dd=Cap(self.fname,[float(self.sMin1.get()),float(self.sMax1.get())])
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[2][0])
        self.result.append(dd[3][0])
        self.result.append("pF")
        line1=np.poly1d(([0,dd[2][0]]))
        x1 = np.linspace(max(dd[0][0],float(self.sMin1.get())),min(dd[0][-1],float(self.sMax1.get())), 100)
        self.ax.clear()
        
        self.ax.plot(dd[0],dd[1],".",x1,line1(x1))
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.ticklabel_format(axis='x',style='sci',scilimits=(0,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("Capacitance [pF]",loc='top')
        self.ax.text(0.5,0.2, ("Result:\n%.4f$\pm$%.4f [pF]" % (dd[2][0],dd[3][0])), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        self.canvas.draw()
        self.toolbar.update()







def Cap(fname,limits):
    f=open(fname)
    d=f.read().splitlines()
    d.pop(0)
    x=[]
    y=[]
    for k in d:
        item=k.split("\t")
        x.append(float(item[0]))
        y.append(float(item[1])*1e12)

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

    
    val1,cov1=np.polyfit(xx[idmin1:idmax1],yy[idmin1:idmax1],0,cov=True)
    err1=np.sqrt(np.diag(cov1))


    return xx,yy,val1,err1

      
