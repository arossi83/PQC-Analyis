import tkinter
from tkinter import *
from tkinter import filedialog,StringVar
from tkinter.ttk import Frame, Button, Style
import tkinter.font as font
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
    
def new_window(Win_class,title,fname,const=1,save=True):
    win2 = Toplevel(root)
    q=Win_class(win2,title,fname,const,save)
    return q.result
    
class mainWindow(object):
    def __init__(self,master):
        self.master=master

        frame = Frame(master, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)

        frame1 = Frame(frame,relief=RIDGE,borderwidth=2)
        frame1.pack(fill=X)
        self.lf=Label(frame1,textvariable=dname, font=( 'bold'), justify=CENTER)
        self.lf.grid(column=0,row=1,columnspan=4,sticky='nesw')#pack(side=LEFT, padx=5, pady=25, expand=True)        
        self.b1=Button(frame1,text="Select Directory",command=self.LoadDir)
        self.b1.grid(column=4,row=1,sticky='nesw')#pack(side=RIGHT, padx=5, expand=True)
        for row in range(3):
            frame1.grid_rowconfigure(row, weight=1) 
        for col in range(5):
            frame1.grid_columnconfigure(col, weight=1)

        frame2 = Frame(frame,relief=RIDGE,borderwidth=2)
        frame2.pack(fill=BOTH)
        self.lt=Label(frame2,text="Select the analysis to be done:",justify=CENTER,font=('Helvetica', 18, 'bold'))
        self.lt.grid(row=0, column=0, columnspan=4, padx=5, pady=5)
        self.chkFlute=[]        
        self.chkList=[]
        ick=0
        self.chkFlute.append(Checkbutton(frame2, text="Flute1", font=('Helvetica',12,'bold'),justify=CENTER, anchor=W, var=ckFlute[0],command=self.SelFlute1Func))
        self.chkFlute[0].grid(column=0, row=1)
        self.chkFlute.append(Checkbutton(frame2, text="Flute2", font=('Helvetica',12,'bold'),justify=CENTER, anchor=W, var=ckFlute[1],command=self.SelFlute2Func))
        self.chkFlute[1].grid(column=1, row=1)
        self.chkFlute.append(Checkbutton(frame2, text="Flute3", font=('Helvetica',12,'bold'),justify=CENTER, anchor=W, var=ckFlute[2],command=self.SelFlute3Func))
        self.chkFlute[2].grid(column=2, row=1)
        self.chkFlute.append(Checkbutton(frame2, text="Flute4", font=('Helvetica',12,'bold'),justify=CENTER, anchor=W, var=ckFlute[3],command=self.SelFlute4Func))
        self.chkFlute[3].grid(column=3, row=1)
        
        for col,fl in enumerate(flute):
            for i in range(nM[fl]):
                self.chkList.append(Checkbutton(frame2, text=mName[fl][i], justify=LEFT, anchor=W, var=ckVar[fl][i]))
                self.chkList[ick].grid(column=col, row=i+2,sticky="nesw")
                ick+=1

        self.chkListA=Checkbutton(frame2, text="All",width=10, var=ckAll,relief=RAISED,font=('Helvetica',12, 'bold'),command=self.SelAllFunc)
        self.chkListA.grid(column=0, row=max(nM.values())+1,columnspan=2)
        subFrame=Frame(frame2,relief=GROOVE,borderwidth=2)
        subFrame.grid(column=0,row=max(nM.values())+2,columnspan=4,sticky='nesw')
        for row in range(max(nM.values())+2):
            frame2.grid_rowconfigure(row, weight=1)
        for col in range(len(flute)):
            frame2.grid_columnconfigure(col, weight=1)        
        #subFrame.pack(fill=BOTH)
        self.chkExtra=Checkbutton(subFrame, text="Standard DiodeCV", var=ckExt,anchor=W,command=self.ExtraStatusFunc)
        self.chkExtra.grid(column=0, row=0)
        self.sl=Label(subFrame,textvariable=extraf, justify=CENTER,state=DISABLED)
        self.sl.grid(column=1,row=0,columnspan=3,sticky='nesw')#pack(side=LEFT, padx=5, pady=25, expand=True)
        self.sb=Button(subFrame,text="Select File",command=self.LoadExtraFile,state=DISABLED)
        self.sb.grid(column=4,row=0,sticky='nesw')#pack(side=RIGHT, padx=5, expand=True)
        subFrame.grid_rowconfigure(0, weight=1)
        for col in range(1,5):
            subFrame.grid_columnconfigure(col, weight=1)            
            
        frame3 = Frame(frame,relief=RIDGE,borderwidth=2)
        frame3.pack(fill=BOTH)
        self.lfo=Label(frame3,text="Output File:", anchor=W, justify=LEFT, width=10)
        self.lfo.pack(side=LEFT, padx=5, pady=15)
        self.fo=Entry(frame3, textvariable=outfname, width=40)
        self.fo.pack(side=LEFT,fill=X,padx=5, expand=True)
        self.outChk=Checkbutton(frame3, text="Save Fig", anchor=W, var=saveFig)
        self.outChk.pack(side=RIGHT)
        self.outChk=Checkbutton(frame3, text="Report", anchor=W, var=createOut)
        self.outChk.pack(side=RIGHT)
        
        self.b3=Button(master,text='Close',command=self.cleanup)
        self.b3.pack(side=RIGHT, padx=5, pady=5)
        self.b2=Button(master,text='Process',command=self.process)
        self.b2.pack(side=RIGHT)

    def SelAllFunc(self):
        temp=ckAll.get()
        for col,fl in enumerate(flute):
            ckFlute[col].set(temp)
            for i in range(nM[fl]):
                ckVar[fl][i].set(temp)

    def SelFlute1Func(self):
        temp=ckFlute[0].get()
        for i in range(nM[flute[0]]):
            ckVar[flute[0]][i].set(temp)

    def SelFlute2Func(self):
        temp=ckFlute[1].get()
        for i in range(nM[flute[1]]):
            ckVar[flute[1]][i].set(temp)
            
    def SelFlute3Func(self):
        temp=ckFlute[2].get()
        for i in range(nM[flute[2]]):
            ckVar[flute[2]][i].set(temp)

    def SelFlute4Func(self):
        temp=ckFlute[3].get()
        for i in range(nM[flute[3]]):
            ckVar[flute[3]][i].set(temp)
            
    def ExtraStatusFunc(self):
        state = str(self.sb['state'])
        if state=='normal':
            self.sl.config(state='disabled')
            self.sb.config(state='disabled')
        else:
            self.sl.config(state='normal')
            self.sb.config(state='normal')
                        
    def process(self):
        if dname.get()=="Please select a directory" or dname.get()=="":
            tkinter.messagebox.showwarning(title=None, message="Please Select a Directory!")
        elif createOut.get() and outfname.get()=="":
            tkinter.messagebox.showwarning(title=None, message="Please Insert An Output File Name!")
        else:
            Res=[]
            onlyfiles = [dname.get()+"/"+f for f in os.listdir(dname.get())]

            for fl in flute:
                for i in range(nM[fl]):
                    indices=[]
                    if ckVar[fl][i].get():
                        #print(("%s - %d/%d - %s - %s" % (fl,i,nM[fl],mTag[fl][i],mType[fl][i])))
                        for fidx,fstr in enumerate(onlyfiles):
                            for tag in mTag[fl][i]:
                                if tag in fstr and fl in fstr and ".txt" in fstr:
                                    indices.append(fidx)
                                    break;
                        if len(indices)>1:
                            sel=SelFileWindow(WinSel,onlyfiles,indices)
                            fname=onlyfiles[indices[sel]]
                        else:
                            fname=onlyfiles[indices[0]]
                        print(("Analyze %s  data" % (mTag[fl][i][0])))
                        if mType[fl][i]=="C":
                            Res.append(new_window(WinCap,mName[fl][i],fname,1,saveFig.get()))
                        elif mType[fl][i]=="CV":
                            Res.append(new_window(WinDiodeCV,mName[fl][i],fname,1,saveFig.get()))
                        elif mType[fl][i]=="D":
                            Res.append(new_window(WinDielBreak,mName[fl][i],fname,1,saveFig.get()))
                        elif mType[fl][i]=="F":
                            Res.append(new_window(WinFET,mName[fl][i],fname,1,saveFig.get()))
                        elif mType[fl][i]=="G":
                            Res.append(new_window(WinGCD,mName[fl][i],fname,1,saveFig.get()))
                        elif mType[fl][i]=="IV":
                            Res.append(new_window(WinDiodeIV,mName[fl][i],fname,1,saveFig.get()))
                        elif mType[fl][i]=="M":
                            Res.append(new_window(WinMos,mName[fl][i],fname,1,saveFig.get()))
                        else:
                            Res.append(new_window(WinRes,mName[fl][i],fname,mType[fl][i],saveFig.get()))
                        print(("%s done" % (mTag[fl][i][0])))
                    else:
                        Res.append([0,0,"SKIPPED"])
            if ckExt.get():
                print("Analyze Standard DiodeCV data")
                Res.append(new_window(WinDiodeCV,"Standard Diode C/V",dname.get()+"/"+extraf.get(),1,saveFig.get()))
                print("Standard DiodeCV done")
                    
            if createOut.get():
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
                worksheet.write_formula('E20', '=(4*B4*13*13/3/33/33)*(1+13/2/(33-13))')
                worksheet.write_formula('E22', '=(4*B3*13*13/3/33/33)*(1+13/2/(33-13))')
                worksheet.write_formula('F20', '=(4*C4*13*13/3/33/33)*(1+13/2/(33-13))')
                worksheet.write_formula('F22', '=(4*C5*13*13/3/33/33)*(1+13/2/(33-13))')
                worksheet.write_formula('G20','=B20-E20')
                worksheet.write_formula('H20','=SQRT(C20*C20+F20*F20)')
                worksheet.write_formula('G22','=B22-E22')
                worksheet.write_formula('H22','=SQRT(C22*C22+E22*E122)')
                worksheet.write_formula('G18','=128.5*B17/B18')
                worksheet.write_formula('H18','=G18*SQRT(C17*C17/(B17*B17)+C18*C18/(B18*B18))')
                worksheet.write_formula('G8','=128.5*B4/B8')
                worksheet.write_formula('H8','=G8*SQRT(C4*C4/(B4*B4)+C8*C8/(B8*B8))')
                worksheet.write_formula('G11','=128.5*B5/B11')
                worksheet.write_formula('H11','=G11*SQRT(C5*C5/(B5*B5)+C11*C11/(B11*B11))')
                worksheet.write_formula('G10','=B10*0.000000000001/1.6E-19/5415000000/0.00505')
                worksheet.write_formula('H10','=C10*0.000000000001/1.6E-19/5415000000/0.00505')
                worksheet.write_formula('G21','=B21*0.000000000001/1.6E-19/5415000000/0.00723')
                worksheet.write_formula('H21','=C21*0.000000000001/1.6E-19/5415000000/0.00723')

                workbook.close()

        
    def cleanup(self):
        self.master.destroy()
        
    def LoadDir(self):   
        dname.set(filedialog.askdirectory(initialdir=mainDir, mustexist=TRUE))

    def LoadExtraFile(self):
        if dname.get()=="Please select a directory" or dname.get()=="":
            tkinter.messagebox.showwarning(title=None, message="Please Select a Directory!")
        else:   
            extraf.set(filedialog.askopenfilename(initialdir=dname.get(), filetypes = (("text files","*.txt"),("all files","*.*"))).split("/")[-1])

        
if __name__ == "__main__":
    root=Tk()
    root.title("PQC Analysis")

    dname=StringVar()
    dname.set("Please select a directory")
    extraf=StringVar()
    extraf.set("Please select a file")

    flute=["flute1","flute2","flute3","flute4"]
    
    nM={"flute1": 6,
        "flute2": 5,
        "flute3": 7,
        "flute4": 6}
    
    mName={"flute1":
           ["Capacitor Measurement",
            "Polysilicon Resitance",
            "n+ Resistance",
            "p-stop Resistance",
            "FET Measurement",
            "MOS Measurement"],
           "flute2":
           ["n+Linewidth Resistance",
            "Polysilicon Meander Resistance",
            "Gated Diode Flute2",
            "pstopLinewidth Resistance",
            "Dielectric Breakdown"],
           "flute3":
           ["BulkCross Resistance",
            "Diode I/V",
            "Diode C/V Depletion Voltage",
            "MetalCover Resistance",
            "p+ Cross Resistance",
            "p+Bridge Resistance",
            "MetalMeanderChain Resistance"],
           "flute4":
           ["n+CBKR Resistance",
            "Gated Diode Flute4",
            "polyCBKR Resistance",
            "PolyChain Resistance",
            "p+Chain Resistance",
            "n+Chain Resistance"]}

    mTag={"flute1": [["Capacitor"],["Poly"],["n+"],["pstop"],["FET"],["MOS"]],
           "flute2": [["n+_linewidth"],["PolyMeander"],["GCD"],["pstopLinewidth"],["DielectricBreak"]],
           "flute3": [["BulckCross"],["DiodeIV","Diode_IV"],["DiodeCV","Diode_CV"],["MetalCover"],["p+Cross"],["p+Bridge"],["Metal_Meander_Chain"]],
           "flute4": [["n+CBKR"],["GCD"],["polyCBKR"],["Poly_Chain"],["p+_Chain"],["n+_Chain"]]}


    
    mType={"flute1":
           ["C",4.53,4.53,4.53,"F","M"],
           "flute2":
           [1,1,"G",1,"D"],
           "flute3":
           [10.726*0.0187*1.218,"IV","CV",4.53,4.53,1,1],
           "flute4":
           [1,"G",1,1,1,1]}

    
    ckVar={"flute1": [],
           "flute2": [],
           "flute3": [],
           "flute4": []}
    ckFlute=[]
    ckExt=BooleanVar()
    ckExt.set(False)
    for col,fl in enumerate(flute):
        ckFlute.append(BooleanVar())
        ckFlute[col].set(True)
        for i in range(nM[fl]):
            ckVar[fl].append(BooleanVar())
            ckVar[fl][i].set(True)


    ckAll=BooleanVar()
    ckAll.set(True)

    createOut=BooleanVar()
    createOut.set(True)
    saveFig=BooleanVar()
    saveFig.set(True)
    outfname=StringVar()
    outfname.set("")

    m=mainWindow(root)
    root.after(100, lambda: root.focus_force())
    root.mainloop()

    
