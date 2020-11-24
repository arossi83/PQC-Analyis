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



class WinDiodeCV:
    def __init__(self, root,title,fname,const):
        self.root = root
        self.root.geometry("+150+50")
        self.fname =fname
        self.title=title
        self.const=const
        self.root.wm_title(self.title)
        self.sMin1=StringVar()
        self.sMin1.set("80")
        self.sMax1=StringVar()
        self.sMax1.set("220")
        self.sMin2=StringVar()
        self.sMin2.set("300")
        self.sMax2=StringVar()
        self.sMax2.set("400")
        
        frame = Frame(self.root, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        frameD1 = Frame(self.root)
        frameD1.pack(fill=BOTH, expand=True)
        frameD2 = Frame(self.root)
        frameD2.pack(fill=BOTH, expand=True)
        
        self.frame1 = Frame(frame)
        self.frame1.pack(fill=X)
        self.fig = Figure(figsize=(7, 6), dpi=100)

        dd=diodecv(self.fname,[float(self.sMin1.get()),float(self.sMax1.get()),float(self.sMin2.get()),float(self.sMax2.get())])

        if "DiodeCV" in fname and "flute" in fname:
            b=fname.split("DiodeCV")[1].split(".t")[0]
        elif "Diode_CV" in fname and "flute" in fname:
            b=fname.split("Diode_CV")[1].split(".t")[0]
        elif "half" in fname and not "flute" in fname:
            b=fname.split("CV")[1].split(".t")[0]

        dirD=os.path.dirname(os.path.abspath(fname))
        if "half" in fname and not "flute" in fname:
            self.img=dirD+"/Diode_half_CV_img"+b+".pdf"
        else:
            self.img=dirD+"/DiodeCV_flute3_img"+b+".pdf"
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[6])
        self.result.append(dd[7])
        self.result.append("")
        
        line1=np.poly1d(dd[2])
        #x1 = np.linspace(float(self.sMin1.get()),float(self.sMax1.get()), 100)
        line2=np.poly1d(dd[3])
        #x2 = np.linspace(float(self.sMin2.get()),float(self.sMax2.get()), 100)
        x1_l = np.linspace(float(self.sMin1.get()),dd[6]*1.02, 100)
        x2_l = np.linspace(dd[6]*0.9,float(self.sMax2.get()), 100)
        
        self.ax=self.fig.add_subplot(111)
        self.ax.plot(dd[0],dd[1],".",x1_l,line1(x1_l),x2_l,line2(x2_l))
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("1/Capacitance$^{2}$ [1/pF$^{2}$]",loc='top')
        self.ax.text(0.7,0.2, ("Result:\n%.2f$\pm$%.2f [V]" % (dd[6],dd[7])), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        
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
        self.b2=Button(self.root,text='Update',command= self.UpdateDiodeCV)
        self.b2.pack(side=RIGHT)

        self.root.wait_window(self.root)
        
    def closeExt(self):
        self.fig.savefig(self.img)
        self.root.destroy()

    def UpdateDiodeCV(self):
        dd=diodecv(self.fname,[float(self.sMin1.get()),float(self.sMax1.get()),float(self.sMin2.get()),float(self.sMax2.get())])
        self.result=[]
        self.result.append(self.title)
        self.result.append(dd[6])
        self.result.append(dd[7])
        self.result.append("")
        line1=np.poly1d(dd[2])
        #x1 = np.linspace(float(self.sMin1.get()),float(self.sMax1.get()), 100)
        line2=np.poly1d(dd[3])
        #x2 = np.linspace(float(self.sMin2.get()),float(self.sMax2.get()), 100)
        x1_l = np.linspace(float(self.sMin1.get()),dd[6]*1.02, 100)
        x2_l = np.linspace(dd[6]*0.9,float(self.sMax2.get()), 100)

        self.ax.clear()
        
        self.ax.plot(dd[0],dd[1],".",x1_l,line1(x1_l),x2_l,line2(x2_l))
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [V]",loc='right')
        self.ax.set_ylabel("1/Capacitance$^{2}$ [1/pF$^{2}$]",loc='top')
        self.ax.text(0.7,0.2, ("Result:\n%.2f$\pm$%.2f [V]" % (dd[6],dd[7])), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        self.canvas.draw()
        self.toolbar.update()







def diodecv(fname,limits):
    f=open(fname)
    d=f.read().splitlines()
    d.pop(0)
    x=[]
    y=[]
    for k in d:
        item=k.split("\t")
        x.append(-1*float(item[0]))
        y.append(1./np.power(float(item[1])*1e12,2))

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

    
    val1,cov1=np.polyfit(xx[idmin1:idmax1],yy[idmin1:idmax1],1,cov=True)
    err1=np.sqrt(np.diag(cov1))
    val2,cov2=np.polyfit(xx[idmin2:idmax2],yy[idmin2:idmax2],1,cov=True)
    err2=np.sqrt(np.diag(cov2))
    
    #    line1=np.poly1d(([0,val1[0]]))
    #    x1 = np.linspace(min1,max1, 100)
    #    line2=np.poly1d(([0,val2[0]]))
    #    x2 = np.linspace(min2,max2, 100)

    res=(val2[1]-val1[1])/(val1[0]-val2[0])
    e_res=np.sqrt(np.power(err2[1]/(val1[0]-val2[0]),2)+np.power(err1[1]/(val1[0]-val2[0]),2)+np.power(err1[0]*(val2[1]-val1[1])/np.power((val1[0]-val2[0]),2),2)+np.power(err2[0]*(val2[1]-val1[1])/np.power((val1[0]-val2[0]),2),2))

    #    print(("Results: %.2f+-%.2f" % (res,e_res)))
    
    #    print(("Results : Val1=%.2f+-%.2f - Val2=%.2f+-%.2f - Diff=%.2f+-%.2f" % (val1[0],err1[0],val2[0],err2[0],val2[0]-val1[0],np.sqrt(np.power(err1,2)+np.power(err2,2))[0])))

    return xx,yy,val1,val2,err1,err2,res,e_res

      
