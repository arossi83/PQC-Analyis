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

class WinGCD:
    def __init__(self, root,title,fname,const):
        self.root = root
        self.root.geometry("+150+50")
        self.fname =fname
        self.const=const
        self.title=title
        self.root.wm_title(self.title)
        self.sMin1=StringVar()
        self.sMin1.set("-8")
        self.sMax1=StringVar()
        self.sMax1.set("-2.5")
        self.sMin2=StringVar()
        self.sMin2.set("-1")
        self.sMax2=StringVar()
        self.sMax2.set("4")
        
        frame = Frame(self.root, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        frameD1 = Frame(self.root)
        frameD1.pack(fill=BOTH, expand=True)
        frameD2 = Frame(self.root)
        frameD2.pack(fill=BOTH, expand=True)
        
        self.frame1 = Frame(frame)
        self.frame1.pack(fill=X)
        self.fig = Figure(figsize=(7, 6), dpi=100)

        dd=gcd(self.fname,[float(self.sMin1.get()),float(self.sMax1.get()),float(self.sMin2.get()),float(self.sMax2.get())])

        b=fname.split("GCD")[1].split(".t")[0]
        dirD=os.path.dirname(os.path.abspath(fname))
        if 'flute2' in fname:
            self.img=dirD+"/GCD_flute2_img"+b+".pdf"
        else:
            self.img=dirD+"/GCD_flute4_img"+b+".pdf"
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[2]-dd[3][0])
        self.result.append(np.sqrt(np.power(dd[4],2)+np.power(dd[5][0],2)))
        self.result.append("pA")
        
        line1=np.poly1d(([0,dd[2]]))
        x1 = np.linspace(float(self.sMin1.get()),float(self.sMax1.get()), 100)
        line2=np.poly1d(([0,dd[3][0]]))
        x2 = np.linspace(float(self.sMin2.get()),float(self.sMax2.get()), 100)
        
        self.ax=self.fig.add_subplot(111)
        self.ax.plot(dd[0],dd[1],".",x1,line1(x1),x2,line2(x2))
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("Current [pA]",loc='top')
        self.ax.set_ylim(0,dd[2]*1.2)
        self.ax.text(0.85,0.7, ("Results:\n%.3f$\pm$%.3f [pA]\n%.3f$\pm$%.3f [pA]\n%.3f$\pm$%.3f [pA]" % (dd[2],dd[4],dd[3][0],dd[5][0],dd[2]-dd[3][0],np.sqrt(np.power(dd[4],2)+np.power(dd[5][0],2)))), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        
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
        self.em2=Entry(frameD1, textvariable=self.sMin2, width=10)
        self.em2.pack(side=RIGHT)#, padx=5, pady=5)
        self.lem2=Label(frameD1,text="Min2:", anchor=W,justify=LEFT, width=6)
        self.lem2.pack(side=RIGHT, padx=5, pady=5)

        self.leM1=Label(frameD2,text="Max1:", anchor=W,justify=LEFT, width=6)
        self.leM1.pack(side=LEFT, padx=5, pady=5)
        self.eM1=Entry(frameD2, textvariable=self.sMax1, width=10)
        self.eM1.pack(side=LEFT)#, padx=5, pady=5)
        self.eM2=Entry(frameD2, textvariable=self.sMax2, width=10)
        self.eM2.pack(side=RIGHT)#, padx=5, pady=5)
        self.leM2=Label(frameD2,text="Max2:", anchor=W,justify=LEFT, width=6)
        self.leM2.pack(side=RIGHT, padx=5, pady=5)
                
        self.b3=Button(self.root,text='Next',command=self.closeExt)
        self.b3.pack(side=RIGHT, padx=5, pady=5)
        self.b2=Button(self.root,text='Update',command= self.UpdateGCD)
        self.b2.pack(side=RIGHT)

        self.root.wait_window(self.root)
        
    def closeExt(self):
        self.fig.savefig(self.img)
        self.root.destroy()

    def UpdateGCD(self):
        dd=gcd(self.fname,[float(self.sMin1.get()),float(self.sMax1.get()),float(self.sMin2.get()),float(self.sMax2.get())])
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[2]-dd[3][0])
        self.result.append(np.sqrt(np.power(dd[4],2)+np.power(dd[5][0],2)))
        self.result.append("pA")
        line1=np.poly1d(dd[2])
        x1 = np.linspace(float(self.sMin1.get()),float(self.sMax1.get()), 100)
        line2=np.poly1d(dd[3])
        x2 = np.linspace(float(self.sMin2.get()),float(self.sMax2.get()), 100)


        self.ax.clear()
        
        self.ax.plot(dd[0],dd[1],".",x1,line1(x1),x2,line2(x2))
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("Current [pA]",loc='top')
        self.ax.set_ylim(0,6.5)
        self.ax.text(0.35,0.1, ("Results:\n%.3f$\pm$%.3f [pA]\n%.3f$\pm$%.3f [pA]\n%.3f$\pm$%.3f [pA]" % (dd[2],dd[4],dd[3][0],dd[5][0],dd[2]-dd[3][0],np.sqrt(np.power(dd[4],2)+np.power(dd[5][0],2)))), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        self.canvas.draw()
        self.toolbar.update()





def gcd(fname,limits):
    f=open(fname)
    d=f.read().splitlines()
    d.pop(0)
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
    min2=limits[2]
    max2=limits[3]


    idmin1=int(np.where(xx==min1)[0])
    idmax1=int(np.where(xx==max1)[0])
    idmin2=int(np.where(xx==min2)[0])
    idmax2=int(np.where(xx==max2)[0])


    val1=np.array(np.amax(yy[idmin1:idmax1]))
    err1=np.array(1e-3)
#    val1,cov1=np.polyfit(xx[idmin1:idmax1],yy[idmin1:idmax1],0,cov=True)
#    err1=np.sqrt(np.diag(cov1))
    val2,cov2=np.polyfit(xx[idmin2:idmax2],yy[idmin2:idmax2],0,cov=True)
    err2=np.sqrt(np.diag(cov2))
    
#    line1=np.poly1d(([0,val1[0]]))
#    x1 = np.linspace(min1,max1, 100)
#    line2=np.poly1d(([0,val2[0]]))
#    x2 = np.linspace(min2,max2, 100)
    
#    print(("Results : Val1=%.3f+-%.3f - Val2=%.3f+-%.3f - Diff=%.3f+-%.3f" % (val1[0],err1[0],val2[0],err2[0],val2[0]-val1[0],np.sqrt(np.power(err1,2)+np.power(err2,2))[0])))

    return xx,yy,val1,val2,err1,err2

      
