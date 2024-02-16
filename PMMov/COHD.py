
import re
from tkinter import *
from tkinter import ttk,filedialog
import pickle
import pandas as pd 
import openpyxl

class Mergeapp:
    def __init__(self, root):
        self.root = root
        self.root.title("COHD")
        self.create_widgets()
    
    def create_widgets(self):
        global master_file,cohd_file,result_file,clicked2,entry1,entry3,entry4,drop2
        master_file = StringVar()
        cohd_file = StringVar()
        result_file = StringVar()
        clicked2 = StringVar()

        try:
            master_file.set(pickle.load(open( "pref.dat", "rb" )))
        except:
            pass

        frm = ttk.Frame(root, padding=10)
        frm.pack(side=LEFT)

        row1 = ttk.Frame(frm)
        ttk.Button(row1, width=25, text="Load Master Sheet:", command=self.input_master).pack(side=LEFT,padx=5)
        entry1 = ttk.Entry(row1,width=40,textvariable=master_file)
        entry1.config(state="readonly")
        row1.pack(side=TOP, padx=5, pady=5)
        entry1.pack(side=RIGHT, expand=YES, fill=X)

        buttonrow3 = ttk.Frame(frm)
        ttk.Button(buttonrow3, text="Load", command=self.loadconcerntrate).pack(side=LEFT,padx=15)
        drop2 = OptionMenu(buttonrow3,clicked2, [])
        drop2.pack(side=LEFT,pady=5)
        buttonrow3.pack(side=TOP,pady=5)

        row3 = ttk.Frame(frm)
        ttk.Button(row3, width=25, text="Load COHD:", command=self.input_cohd).pack(side=LEFT,padx=5)
        entry3 = ttk.Entry(row3, width = 40, textvariable=cohd_file)
        entry3.config(state="readonly")
        row3.pack(side=TOP, padx=5, pady=5)
        entry3.pack(side=RIGHT, expand=YES, fill=X)

        row4 = ttk.Frame(frm)
        ttk.Button(row4, width=25, text="Save results as:", command=self.saveresult).pack(side=LEFT,padx=5)
        entry4 = ttk.Entry(row4,width=40,textvariable=result_file)
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

        root.protocol("WM_DELETE_WINDOW", self.close)


    def input_master(self):
        master_file.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"),("all files",
                                                            "*.*")]))
        entry1.xview_moveto(1)

    def input_cohd(self): 
        cohd_file.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"),("all files",
                                                            "*.*")]))
        entry3.xview_moveto(1)
        
    def saveresult(self):
        result_file.set(filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Excel files", "*.xlsx"),("all files",
                                                            "*.*")]))
        entry4.xview_moveto(1)

    def close(self):
        pickle.dump(master_file.get(),open("pref.dat", "wb"))
        root.destroy()

    def clear(self):
        master_file.set("")
        cohd_file.set("")
        result_file.set("")
        self.log.delete(1.0, END) 


    def logprint(self,text):
        self.log.insert(END, text + '\n')

    def updatedropdown(self):
        global clicked2, drop2
        drop2['menu'].delete(0,'end')
        clicked2.set(sheet_names[0])
        for name in sheet_names:
            drop2['menu'].add_command(label=name, command=lambda name=name: clicked2.set(name))

    def loadconcerntrate(self):
        global sheet_names
        sheet_names = []
        try:
            wb = openpyxl.load_workbook(master_file.get())
            sheet_names = wb.sheetnames

        except:
            self.logprint("cant open the file")
        if not sheet_names:
            self.logprint("No sheets found")
            return
        else:
            self.updatedropdown()
        self.logprint("- Loading Master sheets done !")

    def findmatch(self,value,master):
        row = master[master['[Sample ID]']==value]
        empty_values_mask = row['[PMMoV] \n(gc/ 100mL)'].isna()
        filtered_df = row[~empty_values_mask]
        return filtered_df
    
    def merge_df(self,df, master):

        for index, value in enumerate(df['Sample']):
            if(value.startswith("NT")):
                pattern = r'NT_.*(\d{6})'
                match = re.search(pattern,value)
                if match:
                    date = match.group(1)
                date = int(date)
                print(date)
            if(value.startswith("15")):
                row = self.findmatch(value,master)


                # pull the value from master
                start_time = row['[SampleStartTime]\n(HHMM 24-hr)'].values
                flow_rate = row['[FlowRate] \n(in MGD)]'].values
                pmmov = row['[PMMoV] \n(gc/ 100mL)'].values

                # assign to COHD
                df.loc[index,'SampleStartTime (HHMM 24-hr)'] = start_time[0]
                df.loc[index,'FlowRate (in MGD)'] = flow_rate[0]
                df.loc[index,'PMMoVGeneCopies/100ml'] = pmmov[0]
                df.loc[index,'PCRResultDate (YYMMDD)'] = date

        return df

    def result(self):
        try:
            master_df = pd.read_excel(master_file.get(),clicked2.get())
            cohd_df = pd.read_excel(cohd_file.get())
        except EXCEPTION as e :
            print(f"{e}")   

        output_df = self.merge_df(cohd_df,master_df)
        try:
            output_df.to_excel(result_file.get(),index=False)
        except EXCEPTION as e:
            print(f"{e}")

if __name__ == '__main__':
        root = Tk()
        app = Mergeapp(root)
        root.mainloop()