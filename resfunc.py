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



class WinRes:
    def __init__(self, root,title,fname,const):
        self.root = root
        self.root.geometry("+150+50")
        self.fname =fname
        self.const =const
        self.title = title
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

        dd=Res(self.fname,[float(self.sMin1.get()),float(self.sMax1.get())])

        if 'MetalCover' in fname:
            b=fname.split("MetalCover")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/MetalCover_flute3_VdP_img"+b+".pdf"
        if 'n+' in fname:
            b=fname.split("n+")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/n+_VdP_img"+b+".pdf"
        if 'Poly' in fname:
            b=fname.split("Poly")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/Poly_flute1_VdP_img"+b+".pdf"
        if 'pstop' in fname:
            b=fname.split("pstop")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/pstop_flute1_VdP_img"+b+".pdf"
        if 'PolyMeander' in fname:
            b=fname.split("PolyMeander")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/PolyMenader_flute2_img"+b+".pdf"
        if 'p+Cross' in fname:
            b=fname.split("p+Cross")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/p+Cross_flute3_VdP_img"+b+".pdf"
        if 'BulckCross' in fname:
            b=fname.split("BulckCross")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/BulkCross_flute3_VdP_img"+b+".pdf"
        if 'n+CBKR' in fname:
            b=fname.split("n+CBKR")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/n+CBKR_flute4_VdP_img"+b+".pdf"
        if 'polyCBKR' in fname:
            b=fname.split("polyCBKR")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/polyCBKR_flute4_VdP_img"+b+".pdf"
        if 'p+Bridge' in fname:
            b=fname.split("p+Bridge")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/p+Bridge_flute3_VdP_img"+b+".pdf"
        if 'Poly_Chain' in fname:
            b=fname.split("Poly_Chain")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/PolyChain_flute4_VdP_img"+b+".pdf"
        if 'n+_Chain' in fname:
            b=fname.split("n+_Chain")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/n+Chain_flute4_VdP_img"+b+".pdf"
        if 'p+_Chain' in fname:
            b=fname.split("p+_Chain")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/p+Chain_flute4_VdP_img"+b+".pdf"
        if 'Metal_Meander_Chain' in fname:
            b=fname.split("Metal_Meander_Chain")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/MetalMeanderChain_flute3_VdP_img"+b+".pdf"
        if 'n+_linewidth' in fname:
            b=fname.split("n+_linewidth")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/n+Linewidth_flute2_VdP_img"+b+".pdf"
        if 'pstopLinewidth' in fname:
            b=fname.split("pstopLinewidth")[1].split(".t")[0]
            dirD=os.path.dirname(os.path.abspath(fname))
            self.img=dirD+"/pstopLinewidth_flute2_VdP_img"+b+".pdf"

            
        self.result=[]
        self.result.append(self.title)
        self.result.append(1e-3/dd[2][0]*self.const)
        self.result.append(self.const*1e-3*dd[3][0]/np.power(dd[2][0],2))
        self.result.append(self.const)

        
        line1=np.poly1d(dd[2])
        x1 = np.linspace(dd[0][0],dd[0][-1], 100)

        
        self.ax=self.fig.add_subplot(111)
        self.ax.plot(dd[0],dd[1],".",x1,line1(x1))
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.xaxis.get_offset_text().set_position((1,0.1))
        self.ax.xaxis.get_offset_text().set_rotation(-90)
        self.ax.set_xlabel("Voltage [mV]",loc='right')
        self.ax.set_ylabel("Current [A]",loc='top')
        if "MetalCover" in self.title :
            self.ax.text(0.3,0.7, ("Result:\n%.4f$\pm$%.4f [$\Omega$]" % (1e-3/dd[2][0]*self.const,self.const*1e-3*dd[3][0]/np.power(dd[2][0],2))), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        elif "n+" in self.title and (not "CBKR" in self.title) and (not "Chain" in self.title):
            self.ax.text(0.3,0.7, ("Result:\n%.3f$\pm$%.3f [$\Omega$]" % (1e-3/dd[2][0]*self.const,self.const*1e-3*dd[3][0]/np.power(dd[2][0],2))), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        elif "MetalMeanderChain" in self.title:
            self.ax.text(0.3,0.7, ("Result:\n%.3f$\pm$%.3f [$\Omega$]" % (1e-3/dd[2][0]*self.const,self.const*1e-3*dd[3][0]/np.power(dd[2][0],2))), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        else:            
            self.ax.text(0.3,0.7, ("Result:\n%.1f$\pm$%.1f [$\Omega$]" % (1e-3/dd[2][0]*self.const,self.const*1e-3*dd[3][0]/np.power(dd[2][0],2))), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        
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

        self.leM1=Label(frameD1,text="Max1:", anchor=W,justify=RIGHT, width=6)
        self.leM1.pack(side=LEFT, padx=5, pady=5)
        self.eM1=Entry(frameD1, textvariable=self.sMax1, width=10)
        self.eM1.pack(side=LEFT)#, padx=5, pady=5)
                
        self.b3=Button(self.root,text='Next',command=self.closeExt)
        self.b3.pack(side=RIGHT, padx=5, pady=5)
        self.b2=Button(self.root,text='Update',command= self.UpdateRes)
        self.b2.pack(side=RIGHT)

        self.root.wait_window(self.root)
        
    def closeExt(self):
        self.fig.savefig(self.img)
        self.root.destroy()

    def UpdateRes(self):
        dd=Res(self.fname,[float(self.sMin1.get()),float(self.sMax1.get())])

        self.result=[]
        self.result.append(self.title)
        self.result.append(1e-3/dd[2][0]*self.const)
        self.result.append(self.const*1e-3*dd[3][0]/np.power(dd[2][0],2))
        self.result.append(self.const)
        line1=np.poly1d(dd[2])
        x1 = np.linspace(max(dd[0][0],float(self.sMin1.get())),min(dd[0][-1],float(self.sMax1.get())), 100)
        self.ax.clear()
        
        self.ax.plot(dd[0],dd[1],".",x1,line1(x1))
        self.ax.ticklabel_format(axis='y',style='sci',scilimits=(-1,2),useOffset=True,useMathText=True)
        self.ax.ticklabel_format(axis='x',style='sci',scilimits=(0,2),useOffset=True,useMathText=True)
        self.ax.set_title(self.title)
        self.ax.set_xlabel("Voltage [mV]",loc='right')
        self.ax.set_ylabel("Current [A]",loc='top')
        if "MetalCover" in self.title :
            self.ax.text(0.3,0.7, ("Result:\n%.4f$\pm$%.4f [$\Omega$]" % (1e-3/dd[2][0]*self.const,self.const*1e-3*dd[3][0]/np.power(dd[2][0],2))), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        elif "n+" in self.title and not "CBKR" in self.title and not "linewidth" in self.title:
            self.ax.text(0.3,0.7, ("Result:\n%.3f$\pm$%.3f [$\Omega$]" % (1e-3/dd[2][0]*self.const,self.const*1e-3*dd[3][0]/np.power(dd[2][0],2))), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})
        else:            
            self.ax.text(0.3,0.7, ("Result:\n%.1f$\pm$%.1f [$\Omega$]" % (1e-3/dd[2][0]*self.const,self.const*1e-3*dd[3][0]/np.power(dd[2][0],2))), fontsize=10,horizontalalignment='center', style='normal', transform=self.ax.transAxes, bbox={'facecolor': 'red', 'alpha': 0.5, 'pad': 5})

        self.canvas.draw()
        self.toolbar.update()







def Res(fname,limits):
    f=open(fname)
    d=f.read().splitlines()
    d.pop(0)
    if not 'PolyMeander' in fname:
        d.pop(0)
    x=[]
    y=[]
    for k in d:
        item=k.split("\t")
        x.append(float(item[0])*1e3)
        y.append(float(item[1]))

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

    
    val1,cov1=np.polyfit(xx[idmin1:idmax1],yy[idmin1:idmax1],1,cov=True)
    err1=np.sqrt(np.diag(cov1))


    return xx,yy,val1,err1

      
