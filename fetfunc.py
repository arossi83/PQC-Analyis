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



class WinFET:
    def __init__(self, root,title,fname,const):
        self.root = root
        self.fname =fname
        self.const=const
        self.title=title
        self.root.geometry("+150+50")
        self.root.wm_title(self.title)
        self.sMin1=StringVar()
        self.sMin1.set("-4")
        self.sMax1=StringVar()
        self.sMax1.set("0")
        
        frame = Frame(self.root, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        frameD1 = Frame(self.root)
        frameD1.pack(fill=BOTH, expand=True)
        frameD2 = Frame(self.root)
        frameD2.pack(fill=BOTH, expand=True)
        
        self.frame1 = Frame(frame)
        self.frame1.pack(fill=X)
        self.fig = Figure(figsize=(7, 6), dpi=100)

        dd=FET(self.fname,[float(self.sMin1.get()),float(self.sMax1.get())])

        b=fname.split("FET")[1].split(".t")[0]
        dirD=os.path.dirname(os.path.abspath(fname))        
        self.img=dirD+"/FET_flute1_img"+b+".pdf"
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[4])
        self.result.append(dd[5])
        self.result.append(dd[7])
        
        line1=np.poly1d(dd[3])
        x1 = np.linspace(dd[4]*0.98,dd[4]*1.5, 100)
        line2=np.poly1d(([0,dd[6]]))
        x2 = np.linspace(float(self.sMin1.get()),dd[4]*1.2,100)

        self.ax=self.fig.add_subplot(111)
        self.ax.plot(dd[0],dd[1],".",dd[0][:dd[2].size],dd[2],".",x1,line1(x1),x2,line2(x2),"--")
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("Current [A]",loc='top')
        self.ax.text(0.2,0.5, ("Result:\n%.3f$\pm$%.3f [V]\nExt. V = %.0fmV" % (dd[4],dd[5],1e3*dd[7])), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
    
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.frame1)
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
        self.b2=Button(self.root,text='Update',command= self.UpdateFET)
        self.b2.pack(side=RIGHT)

        self.root.wait_window(self.root)
        
    def closeExt(self):
        self.fig.savefig(self.img)
        self.root.destroy()

    def UpdateFET(self):
        dd=FET(self.fname,[float(self.sMin1.get()),float(self.sMax1.get())])
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[4])
        self.result.append(dd[5])
        self.result.append(dd[7])
        self.ax.clear()        

        line1=np.poly1d(dd[3])
        x1 = np.linspace(dd[4]*0.98,dd[4]*1.5, 100)
        line2=np.poly1d(([0,dd[6]]))
        x2 = np.linspace(float(self.sMin1.get()),dd[4]*1.2,100)

        self.ax.plot(dd[0],dd[1],".",dd[0][:dd[2].size],dd[2],".",x1,line1(x1),x2,line2(x2),"--")
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("Current [pA]",loc='top')
        self.ax.text(0.2,0.5, ("Result:\n%.3f$\pm$%.3f [V]\nExt. V = %.0fmV" % (dd[4],dd[5],1e3*dd[7])), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        self.canvas.draw()
        self.toolbar.update()







def FET(fname,limits):
    f=open(fname)
    d=f.read().splitlines()
    ext=float(d[0].split("\t")[1])
    d.pop(0)
    d.pop(0)
    x=[]
    y=[]
    drl=[]
    for i,k in enumerate(d):
        item=k.split("\t")
        x.append(float(item[0]))
        y.append(float(item[1]))
        if i>0:
            drl.append((y[i]-y[i-1])/(x[i]-x[i-1]))


    m=[]
    m.append(max(drl))
    idmax=drl.index(m[0])
    m.append(y[idmax]-m[0]*x[idmax])
    
    xx=np.array(x)
    yy=np.array(y)
    drp=np.array(drl)
    dr=np.array(m)


    min1=limits[0]
    max1=limits[1]

    idmin1=int(np.where(xx==min1)[0])
    idmax1=int(np.where(xx==max1)[0])
    
    val1,cov1=np.polyfit(xx[idmin1:idmax1],yy[idmin1:idmax1],0,cov=True)
    err1=np.sqrt(np.diag(cov1))
    
    #res=-1*m[1]/m[0]-ext/2
    res=(val1[0]-m[1])/m[0]-ext/2
    e_res=res*0.02

    #print(("Baseline : %f\pm%f" % (val1[0],err1[0])))
    
    return xx,yy,drp,dr,res,e_res,val1[0],ext

      
