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
from mosfunc import WinMos
from fetfunc import WinFET
from gcdfunc import WinGCD
from resfunc import WinRes
from capacitorfunc import WinCap
from dielbreakfunc import WinDielBreak
from diodeCVfunc import WinDiodeCV
from diodeIVfunc import WinDiodeIV
from selfilewin import WinSel


mainDir="/Users/alessandro/Documents/CMS/PQC_Analysis"

def SelFileWindow(Win_class,names,indices):
    win2 = Toplevel(root)
    q=Win_class(win2,names,indices)
    return q.val.get()
    
def new_window(Win_class,title,fname,const=1):
    win2 = Toplevel(root)
    q=Win_class(win2,title,fname,const)
    return q.result
    
class mainWindow(object):
    def __init__(self,master):
        self.master=master

        frame = Frame(master, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)

        frame1 = Frame(frame)
        frame1.pack(fill=X)
        self.lf=Label(frame1,textvariable=dname, anchor=W, justify=LEFT)
        self.lf.pack(side=LEFT, padx=5, pady=25, expand=True)        
        self.b1=Button(frame1,text="Select Directory",command=self.LoadDir)
        self.b1.pack(side=RIGHT, padx=5, expand=True)

        frame3 = Frame(frame)
        frame3.pack(fill=X)
        self.lfo=Label(frame3,text="Output File:", anchor=W, justify=LEFT, width=10)
        self.lfo.pack(side=LEFT, padx=5, pady=5)
        self.fo=Entry(frame3, textvariable=outfname, width=40)
        self.fo.pack(fill=X, padx=5, expand=True)
        
        self.b3=Button(master,text='Close',command=self.cleanup)
        self.b3.pack(side=RIGHT, padx=5, pady=5)
        self.b2=Button(master,text='Process',command=self.process)
        self.b2.pack(side=RIGHT)

    def process(self):
        Res=[]
        onlyfiles = [dname.get()+"/"+f for f in os.listdir(dname.get())]
        indices = [i for i, s in enumerate(onlyfiles) if 'Poly' in s and not ('Chain' in s or 'Meander' in s) and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze Poly data")
        Res.append(new_window(WinRes,"Polysilicon Resitance",fname,4.53))
        print("Poly done")
        indices = [i for i, s in enumerate(onlyfiles) if 'pstop' in s and 'flute1' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze pStop data")
        Res.append(new_window(WinRes,"p-stop Resistance",fname,4.53))
        print("pStop done")
        indices = [i for i, s in enumerate(onlyfiles) if 'n+' in s and 'flute1' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze n+ data")
        Res.append(new_window(WinRes,"n+ Resistance",fname,4.53))
        print("n+ done")
        indices = [i for i, s in enumerate(onlyfiles) if 'MetalCover' in s and 'flute3' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze MetalCover data")
        Res.append(new_window(WinRes,"MetalCover Resistance",fname,4.53))
        print("MetalCover done")
        indices = [i for i, s in enumerate(onlyfiles) if 'PolyMeander' in s and 'flute2' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze PolyMeander data")
        Res.append(new_window(WinRes,"Polysilicon Meander Resistance",fname))
        print("PolyMeander done")
        indices = [i for i, s in enumerate(onlyfiles) if 'p+Cross' in s and 'flute3' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze p+Cross data")
        Res.append(new_window(WinRes,"p+ Cross Resistance",fname,4.53))
        print("p+Cross done")
        indices = [i for i, s in enumerate(onlyfiles) if 'BulckCross' in s and 'flute3' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze BulkCross data")
        Res.append(new_window(WinRes,"BulkCross Resistance",fname,10.726*0.0187*1.218))
        print("BulkCross done")
        indices = [i for i, s in enumerate(onlyfiles) if 'n+CBKR' in s and 'flute4' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze n+CBKR data")
        Res.append(new_window(WinRes,"n+CBKR Resistance",fname))
        print("n+CBKR done")
        indices = [i for i, s in enumerate(onlyfiles) if 'polyCBKR' in s and 'flute4' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze polyCBKR data")
        Res.append(new_window(WinRes,"polyCBKR Resistance",fname))
        print("polyCBKR done")
        indices = [i for i, s in enumerate(onlyfiles) if 'Poly_Chain' in s and 'flute4' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze PolyChain data")
        Res.append(new_window(WinRes,"PolyChain Resistance",fname))
        print("PolyChain done")
        indices = [i for i, s in enumerate(onlyfiles) if 'n+_Chain' in s and 'flute4' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze n+Chain data")
        Res.append(new_window(WinRes,"n+Chain Resistance",fname))
        print("n+Chain done")
        indices = [i for i, s in enumerate(onlyfiles) if 'p+_Chain' in s and 'flute4' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze p+Chain data")
        Res.append(new_window(WinRes,"p+Chain Resistance",fname))
        print("p+Chain done")
        indices = [i for i, s in enumerate(onlyfiles) if 'Metal_Meander_Chain' in s and 'flute3' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze MetalMeanderChain data")
        Res.append(new_window(WinRes,"MetalMeanderChain Resistance",fname))
        print("MetalMeanderChain done")
        indices = [i for i, s in enumerate(onlyfiles) if 'p+Bridge' in s and 'flute3' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze p+Bridge data")
        Res.append(new_window(WinRes,"p+Bridge Resistance",fname))
        print("p+Bridge done")
        indices = [i for i, s in enumerate(onlyfiles) if 'n+_linewidth' in s and 'flute2' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze n+Linewidth data")
        Res.append(new_window(WinRes,"n+Linewidth Resistance",fname))
        print("n+Linewidth done")
        indices = [i for i, s in enumerate(onlyfiles) if 'pstopLinewidth' in s and 'flute2' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze pstopLinewidth data")
        Res.append(new_window(WinRes,"pstopLinewidth Resistance",fname))
        print("pstopLinewidth done")
        indices = [i for i, s in enumerate(onlyfiles) if 'Capacitor' in s and 'flute1' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze Capacitor data")
        Res.append(new_window(WinCap,"Capacitor Measurement",fname))
        print("Capacitor done")
        indices = [i for i, s in enumerate(onlyfiles) if 'FET' in s and 'flute1' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        print("Analyze FET data")
        Res.append(new_window(WinFET,"FET Measurement",fname))
        print("FET done")
        print("Analyze MOS data")
        indices = [i for i, s in enumerate(onlyfiles) if 'MOS' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        Res.append(new_window(WinMos,"MOS Measurement",fname))
        print("MOS done")
        print("Analyze GCD flute2 data")
        indices = [i for i, s in enumerate(onlyfiles) if 'GCD' in s and 'flute2' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        Res.append(new_window(WinGCD,"Gated Diode Flute2",fname))
        print("GCD flute2 done")
        print("Analyze GCD flute4 data")
        indices = [i for i, s in enumerate(onlyfiles) if 'GCD' in s and 'flute4' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        Res.append(new_window(WinGCD,"Gated Diode Flute4",fname))
        print("GCD flute4 done")
        print("Analyze DiodeCV data")
        indices = [i for i, s in enumerate(onlyfiles) if ('DiodeCV' in s or 'Diode_CV' in s) and 'flute3' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        Res.append(new_window(WinDiodeCV,"Diode C/V Depletion Voltage",fname))
        print("DiodeCV done")
        print("Analyze DiodeIV data")
        indices = [i for i, s in enumerate(onlyfiles) if ('DiodeIV' in s or 'Diode_IV' in s) and 'flute3' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        Res.append(new_window(WinDiodeIV,"Diode I/V",fname))
        print("DiodeIV done")
        print("Analyze Dielectric Break data")
        indices = [i for i, s in enumerate(onlyfiles) if 'DielectricBreak' in s and 'flute2' in s and '.txt' in s ]
        if len(indices)>1:
            sel=SelFileWindow(WinSel,onlyfiles,indices)
            fname=onlyfiles[indices[sel]]
        else:
            fname=onlyfiles[indices[0]]
        Res.append(new_window(WinDielBreak,"Dielectric Breakdown",fname))
        print("Dielectric Break done")

        output=dname.get()+"/"+outfname.get()+".xlsx"

        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()
        tformat= workbook.add_format({'bold': True})
        tformat.set_align('center')
        worksheet.set_column('A:A', 40)
        worksheet.set_column('B:C', 20)
        worksheet.set_column('D:H', 15)
        fieldnames = ['Measurement', 'Value', 'Error','ExtraInfo','Correction Factor','C.F. error','Derived Value','Der. Value error']
        for i,data in enumerate(fieldnames):
            worksheet.write(0,i,data,tformat)
        for row_num, row_data in enumerate(Res):
            for col_num, col_data in enumerate(row_data):
                worksheet.write(row_num+1, col_num, col_data)
        #FORMULAS
        worksheet.write_formula('E9', '=(4*B4*13*13/3/33/33)*(1+13/2/(33-13))')
        worksheet.write_formula('E10', '=(4*B2*13*13/3/33/33)*(1+13/2/(33-13))')
        worksheet.write_formula('F9', '=(4*C4*13*13/3/33/33)*(1+13/2/(33-13))')
        worksheet.write_formula('F10', '=(4*C2*13*13/3/33/33)*(1+13/2/(33-13))')
        worksheet.write_formula('G9','=B9-E9')
        worksheet.write_formula('H9','=SQRT(C9*C9+F9*F9)')
        worksheet.write_formula('G10','=B10-E10')
        worksheet.write_formula('H10','=SQRT(C10*C10+E10*E10)')
        worksheet.write_formula('G15','=128.5*B7/B15')
        worksheet.write_formula('H15','=G15*SQRT(C7*C7/(B7*B7)+C15*C15/(B15*B15))')
        worksheet.write_formula('G16','=128.5*B4/B16')
        worksheet.write_formula('H16','=G16*SQRT(C4*C4/(B4*B4)+C16*C16/(B16*B16))')
        worksheet.write_formula('G17','=128.5*B3/B17')
        worksheet.write_formula('H17','=G17*SQRT(C3*C3/(B3*B3)+C17*C17/(B17*B17))')
        worksheet.write_formula('G21','=B21*0.000000000001/1.6E-19/5415000000/0.00505')
        worksheet.write_formula('H21','=C21*0.000000000001/1.6E-19/5415000000/0.00505')
        worksheet.write_formula('G22','=B22*0.000000000001/1.6E-19/5415000000/0.00723')
        worksheet.write_formula('H22','=C22*0.000000000001/1.6E-19/5415000000/0.00723')

        workbook.close()

        
    def cleanup(self):
        self.master.destroy()
        
    def LoadDir(self):   
        dname.set(filedialog.askdirectory(initialdir=mainDir, mustexist=TRUE))

        
if __name__ == "__main__":
    root=Tk()
    root.title("PQC Analysis")

    dname=StringVar()
    dname.set("Please select a directory")
    outfname=StringVar()
    outfname.set("")

    m=mainWindow(root)
    root.after(100, lambda: root.focus_force())
    root.mainloop()

    
