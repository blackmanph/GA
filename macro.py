
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
import ast

class Macro:
    SHEET_COL = ["CP/ul","Initial Volume analyzed (mL)","Final Concentrate Volume (mL)","Volume used for Extraction (ml)"
                 ,"Final RNA Extraction Volume (ul)","CP/100 ml of Sample","Detection Limit (CP/100ml)","Marker Detected?","Droplet QC Pass"]
    DDPCR_COL = ["Sample","PCRType","TargetDetected"
                 ,"DetectedNotQuantifiable","QualityControlPassed","ControlQCPass?","DropletQCPass","DetectionLowerLimit"
                 ,"N1GeneCopies","N2GeneCopies","EGeneCopies","Phi6GeneCopies"
                 ,"Comments","SampleStartTime (HHMM 24-hr)","PCRResultDate (YYMMDD)","FlowRate (in MGD)","PMMoVGeneCopies/100ml"]
    RSV_COL = ["Sample","PCRType","RSVTargetDetected","SC2TargetDetected","NVG1TargetDetected","NVG2TargetDetected"
               ,"DetectedNotQuantifiable","QualityControlPassed","ControlQCPass?","DropletQCPass","DetectionLowerLimit"
               ,"RSVGeneCopies","SC2GeneCopies","NVG1GeneCopies","NVG2GeneCopies","Phi6GeneCopies"
               ,"Comments","SampleStartTime (HHMM 24-hr)","PCRResultDate (YYMMDD)","FlowRate (in MGD)","PMMoVGeneCopies/100ml"]

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
        raw_input.set(filedialog.askopenfilenames(parent=self.root,filetypes=[("CSV files", "*.csv"),("all files",
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
        paths= list(ast.literal_eval(raw_input.get()))
        master_df = pd.read_excel(master_input.get(),clicked2.get())
        overall_df = pd.DataFrame()
        for file_path in paths:
            try:

                raw_df = pd.read_csv(file_path,index_col=False)
            except EXCEPTION as e :
                print(f"{e}")
        # while True:
        #     raw_df = pd.read_csv(paths,index_col=False)   
            filtered_df = raw_df[raw_df["Target"].apply(lambda x: not x.isnumeric())]
            if "CopiesPer20uLWell" or "AcceptedDroplets"in filtered_df.columns:
                filtered_df = filtered_df.rename(columns={"CopiesPer20uLWell": "Copies/20µLWell","AcceptedDroplets": "Accepted Droplets"})
            if "Sample description 1" in filtered_df.columns:
                filtered_df = filtered_df.rename(columns={"Sample description 1": "Sample"})
            print(filtered_df.columns)
            filtered_df = filtered_df[["Sample","Target","Copies/20µLWell","Accepted Droplets","Positives"]].copy()
                
            print(filtered_df.head)
            overall_df = pd.concat([overall_df,filtered_df],ignore_index=True)

        target_list = overall_df["Target"].unique()
        df_dict = {target: pd.DataFrame(columns=overall_df.columns.tolist()+self.SHEET_COL) for target in target_list}
        print(df_dict.keys())
        for index, rows in overall_df.iterrows():
            if rows["Target"] in df_dict:

                # df_dict[rows["Target"]] = pd.DataFrame(rows)  
                df_dict[rows["Target"]] = pd.concat([df_dict[rows["Target"]], pd.DataFrame([rows])], ignore_index=True)
                # print(df_dict[rows["Target"]])
        # print(df_dict)
        for key in df_dict:
            df_dict[key]['Sample'] = df_dict[key]['Sample'].str.replace(' ', '_')
            df_dict[key]['Sample'] = df_dict[key]['Sample'].str.replace('COV','POS')
            df_dict[key]['Sample'] = df_dict[key]['Sample'].str.replace('PHI','POS')
            df_dict[key]['Sample'] = df_dict[key]['Sample'].str.replace('RSV','POS')
            df_dict[key]['Sample'] = df_dict[key]['Sample'].str.replace('NV','POS')

            for index, row in df_dict[key].iterrows():   
                master_id = row["Sample"]
                if not pd.isna(master_id):
                    matching_row_df2 = master_df[master_df['[Sample ID]'].astype(str).str.contains(str(master_id))]
                if not matching_row_df2.empty:
            # Extract the value from the matching row in df2
                    for indexs in matching_row_df2.index:
                        concern = matching_row_df2.loc[indexs, '[Final Concentrate Volume (mL)]']
                        # dilution = matching_row_df2.loc[indexs, '[Dilution factor]']
                        df_dict[key].loc[index, "Final Concentrate Volume (mL)"] = concern
                        # df_dict[key].loc[index,"Dilution Factor"] = dilution

            df_dict[key].sort_values(by='Sample', inplace=True)
            df_dict[key]["CP/ul"] = df_dict[key]["Copies/20µLWell"]/5
            df_dict[key]["Initial Volume analyzed (mL)"] = 100
            df_dict[key]["Volume used for Extraction (ml)"] = 0.2
            df_dict[key]["Final RNA Extraction Volume (ul)"] = 80
            df_dict[key]["Detection Limit (CP/100ml)"] = (0.6 * df_dict[key]['Final RNA Extraction Volume (ul)']) * ((df_dict[key]['Final Concentrate Volume (mL)'] / df_dict[key]['Volume used for Extraction (ml)']) / df_dict[key]['Initial Volume analyzed (mL)']) * 100
            df_dict[key]["CP/100 ml of Sample"] = np.where(df_dict[key]['Accepted Droplets'] >= 3, (((df_dict[key]['CP/ul'] * df_dict[key]['Final RNA Extraction Volume (ul)']) * (df_dict[key]['Final Concentrate Volume (mL)'] / df_dict[key]['Volume used for Extraction (ml)'])) / df_dict[key]['Initial Volume analyzed (mL)']) * 100, df_dict[key]['Detection Limit (CP/100ml)'])
            df_dict[key]["Marker Detected?"] = np.where((df_dict[key]['Accepted Droplets'] >= 3) & (df_dict[key]['Positives'] >= 8000), 1, 0)
            df_dict[key]["Droplet QC Pass"] = np.where(df_dict[key]['Positives']>=8000,1,0)

            # print(df_dict[key]["Final Concentrate Volume (mL)"])
        with pd.ExcelWriter(output.get(), engine='xlsxwriter') as writer:
            # Iterate through the df_dict and write each DataFrame to a new sheet
            for key, df in df_dict.items():
                df.to_excel(writer, sheet_name=key, index=False)


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
