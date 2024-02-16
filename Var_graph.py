
from tkinter import *
from tkinter import ttk,filedialog,messagebox
import pickle
from pathlib import Path
import pandas as pd 
import numpy as np
from io import StringIO 

class Var_graph:
    def __init__(self, root):
        self.root = root
        self.root.title("Variant grapher")
        self.create_widgets()

    def create_widgets(self):
        global raw_input,result_file,entry1,entry4
        raw_input = StringVar()
        result_file = StringVar()
        
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(side=LEFT)

        row2 = ttk.Frame(frm)
        ttk.Button(row2, width=25, text="Load Raw data:", command=self.inputfile).pack(side=LEFT,padx=5)
        entry1 = ttk.Entry(row2,width=40,textvariable=raw_input)
        entry1.config(state="readonly")
        row2.pack(side=TOP, padx=5, pady=5)
        entry1.pack(side=RIGHT, expand=YES, fill=X)
        entry1.xview_moveto(1)

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

        log = Text(height=20,width=50)
        log.pack(side=RIGHT, padx=10,pady=10)

        self.root.protocol("WM_DELETE_WINDOW", self.close)

    def inputfile(self):
        raw_input.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"),("all files",
                                                            "*.*")]))
        entry1.xview_moveto(1)

    def saveresult(self):
        result_file.set(filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Excel files", "*.xlsx"),("all files",
                                                            "*.*")]))
        entry4.xview_moveto(1)

    def clear(self):
        raw_input.set("")
        result_file.set("")
        self.log.delete(1.0, END) 

    def close(self):
        pickle.dump(raw_input.get(),open("pref.dat", "wb"))
        self.root.destroy()

    def logprint(self,text):
        self.log.insert(END, text + '\n')

    def join_str(self,series):
        return ' '.join(series) 

    def result(self):
        #  sort the dataframe ordered by dates
        #  merge the dataframe to one header
        #  iterate through the data rows 
        try:
            inpath = Path(raw_input.get())
            df = pd.read_excel(inpath,header=2)

        except Exception as e:
            self.logprint("- Error reading input file!")
            print(e)
            return 1

        # print(df.columns)

        df['Date'] = pd.to_datetime(df['Date'], format='%y%m%d', errors='coerce').dt.date
        df.columns = ['' if col.startswith('Unnamed') else col for col in df.columns]

        df_valid_dates = df.dropna(subset=['Date'])
        df_valid_dates.sort_values(by='Date', inplace=True)

        col_index = df.columns.get_loc('Total N1')
        col_index = col_index + 2
        for index, value in df_valid_dates.iterrows():
            n1 = value['N1 GC/100mL']
            flow_rate = value['Flow Rate (MGD)']
            total_n1 = n1/100*(flow_rate*1000000*3785.41)
            df_valid_dates.loc[index,'Total N1'] = total_n1
            df_valid_dates.loc[index, df_valid_dates.columns[col_index]:] *= total_n1/100

        # create the daily report
        same_date_df = df_valid_dates[['Date','Site','Total N1'] + list(df_valid_dates.columns[col_index:])].copy()
        # print(same_date_df.loc[0])

        # merge same dates
        agg_columns = {'Site': self.join_str,'Total N1': 'sum', **dict.fromkeys(same_date_df.columns[3:], 'sum')}
        daily_df_merged = same_date_df.groupby('Date').agg(agg_columns).reset_index()
        daily_df_merged.insert(3,'', None)
        
        # calcaulate new percentage
        col_index = daily_df_merged.columns.get_loc('Total N1')
        col_index = col_index + 2
        for index, value in daily_df_merged.iterrows():
            # print(value)
            daily_df_merged.loc[index, daily_df_merged.columns[col_index:]] /= value['Total N1']

        # print(daily_df_merged.loc[0])
        daily_df_merged.replace(0,np.nan,inplace=True)


        # create the weekly report
        current_week_num = 1
        time_info = daily_df_merged.loc[0,'Date'].isocalendar()
        current_week = time_info[1]

        weekly_df = daily_df_merged.copy()
        weekly_df.insert(3,'Total N1 per week', None)
        weekly_df.insert(4,'Week Number', None)
        col_index = weekly_df.columns.get_loc('Week Number')
        col_index = col_index + 2
        # print(weekly_df.loc[0])

        for index, value in weekly_df.iterrows():
            value_info = value['Date'].isocalendar()
            if(value_info[2]==7):
                value_week = value_info[1] + 1
                if(value_week == 53):
                    value_week = 0
            else:
                value_week = value_info[1]
            if value_week != current_week:
                if value_week < current_week:
                    # Handling change to the next year
                    current_week_num += (52 - current_week) + value_week
                elif value_week - current_week > 1:
                    # Handling skipped weeks
                    current_week_num += (value_week - current_week)
                else:
                    current_week_num += value_week - current_week
            current_week = value_week
            weekly_df.loc[index,'Week Number'] = current_week_num
            weekly_df.loc[index, weekly_df.columns[col_index:]] *= value['Total N1']
            # print(weekly_df.loc[index])

        weekly_df.replace(0,np.nan,inplace=True)
        agg_columns = {'Date': lambda x: x.iloc[0],'Site': self.join_str,'Total N1': 'sum', **dict.fromkeys(weekly_df.columns[6:], 'sum')}

        weekly_df_merged = weekly_df.groupby('Week Number').agg(agg_columns).reset_index()
        # print(weekly_df_merged.loc[0])
        weekly_df_merged.insert(4,'',None)
        col_index = weekly_df_merged.columns.get_loc('Total N1')
        col_index = col_index + 2

        for index, value in weekly_df_merged.iterrows():
            # print(value)
            weekly_df_merged.loc[index, weekly_df_merged.columns[col_index:]] /= value['Total N1']

        try:
            with pd.ExcelWriter(result_file.get()) as writer:
                
                df_valid_dates.to_excel(writer,sheet_name='Perday_count', index=False)
                daily_df_merged.to_excel(writer,sheet_name='Perday_percent',index=False)
                weekly_df.to_excel(writer,sheet_name='Perweek_count',index=False)
                weekly_df_merged.to_excel(writer,sheet_name='Perweek_percent',index=False)


        except EXCEPTION as e:
            print(f"{e}")

if __name__ == '__main__':
    root = Tk()
    app = Var_graph(root)
    root.mainloop()

