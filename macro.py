
from tkinter import *
from tkinter import ttk,filedialog,messagebox
import pickle,os,sys,csv
from pathlib import Path
import pandas as pd 
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill
import xlsxwriter
import os
import math
from numbers import Number

class Macro:

    search_word = "Sample Name"
    sheet_name = "Results"
    starting_word = "Well"
    filter_word = ["Sample Name", "Ct Mean","Ct SD"]
    sample =[]

    wb = openpyxl.Workbook()
    red_fill = PatternFill(patternType='solid', fgColor= '00FF0000')


    def __init__(self, root):
        self.root = root
        self.root.title("PMMoV Calculation")

        self.create_widgets()

    def create_widgets(self):
        global raw_input,output,master_input,entry1,entry4,clicked2,drop2
        raw_input = StringVar()
        output = StringVar()
        master_input = StringVar()
        clicked2 = StringVar()
        try:
            raw_input.set(pickle.load(open( "pref.dat", "rb" )))
        except:
            pass

        frm = ttk.Frame(self.root, padding=10)
        frm.pack(side=LEFT)

        row5 = ttk.Frame(frm)
        ttk.Button(row5, width=25, text="Load Master sheet:", command=self.inputmaster).pack(side=LEFT,padx=5)
        entry5 = ttk.Entry(row5,width=40,textvariable=master_input)
        entry5.config(state="readonly")
        row5.pack(side=TOP, padx=5, pady=5)
        entry5.pack(side=RIGHT, expand=YES, fill=X)
        entry5.xview_moveto(1)

        buttonrow3 = ttk.Frame(frm)
        ttk.Button(buttonrow3, text="Load", command=self.loadmaster).pack(side=LEFT,padx=15)
        drop2 = OptionMenu(buttonrow3,clicked2, [])
        drop2.pack(side=LEFT,pady=5)
        buttonrow3.pack(side=TOP,pady=5)

        row1 = ttk.Frame(frm)
        ttk.Button(row1, width=25, text="Load Raw data:", command=self.inputfile).pack(side=LEFT,padx=5)
        entry1 = ttk.Entry(row1,width=40,textvariable=raw_input)
        entry1.config(state="readonly")
        row1.pack(side=TOP, padx=5, pady=5)
        entry1.pack(side=RIGHT, expand=YES, fill=X)

        row4 = ttk.Frame(frm)
        ttk.Button(row4, width=25, text="Save results as:", command=self.saveresult).pack(side=LEFT,padx=5)
        entry4 = ttk.Entry(row4,width=40,textvariable=output)
        entry4.config(state="readonly")
        row4.pack(side=TOP, padx=5, pady=5)
        entry4.pack(side=RIGHT, expand=YES, fill=X)

        buttonrow2 = ttk.Frame(frm)
        ttk.Button(buttonrow2, text="Run", command=self.result).pack(side=LEFT,padx=15)
        ttk.Button(buttonrow2, text="Clear", command=self.clear).pack(side=LEFT,padx=15)
        ttk.Button(buttonrow2, text="Close", command=self.close).pack(side=LEFT,padx=15)
        buttonrow2.pack(side=BOTTOM,pady=5)
        
        self.log = Text(height=20,width=50)
        self.log.pack(side=RIGHT, padx=10,pady=10)

        self.root.protocol("WM_DELETE_WINDOW", self.close)

    def inputmaster(self):
        master_input.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"),("all files",
                                                            "*.*")]))
        
    def inputfile(self):
        raw_input.set(filedialog.askopenfilename(filetypes=[("CVS files", "*.csv"),("all files",
                                                            "*.*")]))
        entry1.xview_moveto(1)
        
    def saveresult(self):
        output.set(filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Excel files", "*.xlsx"),("all files",
                                                            "*.*")]))
        entry4.xview_moveto(1)

    def close(self):
        pickle.dump(raw_input.get(),open("pref.dat", "wb"))
        self.root.destroy()

    def logprint(self,text):
        self.log.insert(END, text + '\n')


    def updatedropdown(self):
        global clicked2, drop2
        drop2['menu'].delete(0,'end')
        clicked2.set(sheet_names[0])
        for name in sheet_names:
            drop2['menu'].add_command(label=name, command=lambda name=name: clicked2.set(name))

    def loadmaster(self):
        global sheet_names, wb
        sheet_names = []
        try:
            wb = openpyxl.load_workbook(master_input.get())
            sheet_names = wb.sheetnames

        except:
            self.logprint("cant open the file")
        if not sheet_names:
            self.logprint("No sheets found")
            return
        else:
            self.updatedropdown()
        self.logprint("- Loading Master sheets done !")


    def output_to_Excel(self,data,name):
        try:

            data.to_excel(name,index=False)
            
        except:
            self.logprint("- Error writing output file!")
            messagebox.showerror('Error', 'Error writing output file!') 
            return 1
        

    def result(self):
        print("hello")
        try:
            master_df = pd.read_excel(master_input.get(),clicked2.get())
            raw_df = pd.read_csv(raw_input.get())
        except EXCEPTION as e :
            print(f"{e}")   
        filtered_df = raw_df[raw_df["Target"].apply(lambda x: not x.isnumeric())]
        filtered_df = filtered_df[["Sample description 1","Target","Copies/20µLWell","Accepted Droplets","Positives"]].copy()
        target_list = filtered_df["Target"].unique()
        print(filtered_df.head)

        df_dict = {target: pd.DataFrame(columns=filtered_df.columns) for target in target_list}
        print(df_dict.keys())
        for index, rows in filtered_df.iterrows():
            if rows["Target"] in df_dict:

                # df_dict[rows["Target"]] = pd.DataFrame(rows)  
                df_dict[rows["Target"]] = pd.concat([df_dict[rows["Target"]], pd.DataFrame([rows])], ignore_index=True)
                print(df_dict[rows["Target"]])
        print(df_dict)


    def clear(self):
        global sheet_names,drop2,constrain_input,constrain_output,machine_str
        raw_input.set("")
        master_input.set("")
        sheet_names = []
        clicked2.set(' ')
        drop2['menu'].delete(0,END)
        self.log.delete(1.0, END) 
        
if __name__ == '__main__':
    root = Tk()
    app = Macro(root)
    root.mainloop()
