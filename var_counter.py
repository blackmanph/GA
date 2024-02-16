from tkinter import *
from tkinter import ttk,filedialog,messagebox
import pickle,os,sys,csv
from pathlib import Path

def openfile():
    dfname.set(filedialog.askopenfilename(filetypes=[('CSV (Comma delimited)', '*.csv')]))
    entry1.xview_moveto(1)

def opendir(): #TODO: show all files
    idname.set(filedialog.askdirectory())
    entry2.xview_moveto(1)
    
def savefile():
    ofname.set(filedialog.asksaveasfilename(defaultextension='.tsv', filetypes=[("TSV (Tab delimited)", '*.tsv')], initialfile="variant_counts.tsv"))
    entry3.xview_moveto(1)

def close():
    pickle.dump(dfname.get(),open("pref.dat", "wb"))
    root.destroy()

def logprint(text):
    log.insert(END, text + '\n')

def run():
    log.delete(1.0,END)
    dictpath = Path(dfname.get())
    inpath = Path(idname.get())
    outpath = Path(ofname.get())
    ### READ IN VARIANT DICTIONARY ###
    logprint("Variant Counter")
    variants_dict = {}

    try:
        with open(dictpath, "r") as dictcsv:
            variants = csv.reader(dictcsv)
            next(variants)
            for line in variants:
                variants_dict[line[0]] = [int(line[1]), line[2].split()]
        logprint("- Successfully read variant dictionary.")
    except:
        logprint("- Error reading variant dictionary!")
        messagebox.showerror('Error', 'Error reading variant dictionary!')
        return 1

    ### PROMPT FOR INPUT DIRECTORY ###
    try:
        ### GET LIST OF VALID INPUT FILES ###
        tsv_files = [
            os.path.join(root, file)
            for root, _, files in os.walk(idname.get())
            for file in files
            if (
                file.lower().endswith("chim_rm.tsv") or
                file.lower().endswith("covar_deconv.tsv")

            ) and "Collected" not in file
            # if file.lower().endswith(".tsv")

        ]
        for tsv_file in tsv_files:
            print("Matching TSV File:", tsv_file)
    except:
        logprint("- Error opening directory!")
        messagebox.showerror('Error', 'Error opening directory!')
        return 1

    if not tsv_files:
        logprint("- No files were found!")
        messagebox.showwarning('Warning', 'No input files found!')
        return 0

    rows = {}
    header = (
        ["Code", "Date", "Number of Reads"]
        + [var for var in variants_dict]
        + ["Mixed", "Mixed variants", "Other", "Other variants"]
    )
    file_names = [os.path.basename(f) for f in tsv_files]
    ids = sorted([*set([name[5:13].replace('_', '') for name in file_names])])
    for i in ids:
        if i is not None:
            rows[i] = {}
            rows[i]["Code"] = i[:2]
            rows[i]["Date"] = i[2:8]
            rows[i]["Number of Reads"] = 0
            for var in variants_dict:
                rows[i][var] = 0
            rows[i]["Mixed"] = 0
            rows[i]["Mixed variants"] = ""
            rows[i]["Other"] = 0
            rows[i]["Other variants"] = ""

            logprint("- Found data on:")
            for date in sorted([*set([i[2:8] for i in ids])]):
                logprint(
                        f"    {date[:2]}.{date[2:4]}.{date[4:]} from {', '.join([i[:2] for i in ids if i[2:] == date])}"
                )

    ### PROCESS DATA ###
    for file in tsv_files:
        try:
            in_file = open(inpath / file, "r")
        except:
            logprint("- Error opening %s!" % file)
            messagebox.showerror('Error', 'Error opening %s!' % file)
            return 1
        id = os.path.basename(file)[5:13].replace('_', '')
        reader = csv.reader(in_file, delimiter="\t")
        next(reader)
        next(reader)
        for line in reader:  # search for matches
            matches = []
            for var in variants_dict:
                check = 0
                for SNP in variants_dict[var][1]:
                    if SNP[0] == "!":
                        if SNP[1:] in line[0]:
                            check += 1
                    elif SNP not in line[0]:
                        check += 1
                if check <= variants_dict[var][0]:
                    matches.append(var)
            if len(matches) > 1:  # if matches multiple variants
                rows[id]["Mixed"] += float(line[1])
                if not "'" + line[0] + "'" in rows[id]["Mixed variants"]:
                    rows[id]["Mixed variants"] += "'" + line[0] + "', "
            elif matches:  # if matches one variant, classify accordingly
                rows[id][matches[0]] += float(line[1])
            else:  # if matches no variants, classify as "other" and record sequence data
                rows[id]["Other"] += float(line[1])
                if not "'" + line[0] + "'" in rows[id]["Other variants"]:
                    rows[id]["Other variants"] += "'" + line[0] + "', "
            rows[id]["Number of Reads"] += float(line[1])
        in_file.close()

    ### WRITE OUTPUT FILES ###
    try:
        with open(outpath, "w", newline='') as f:
            writer = csv.DictWriter(f, fieldnames=header, delimiter="\t")
            writer.writeheader()
            for id in rows:
                for var in [i for i in variants_dict] + ["Mixed", "Other"]:
                    rows[id][var] = round((rows[id][var] / rows[id]["Number of Reads"]), 3)
                writer.writerow(rows[id])
    except:
        logprint("- Error writing output file!")
        messagebox.showerror('Error', 'Error writing output file!') # TODO show output file name
        return 1

    logprint("- Successfully wrote output file") #TODO show output file name
    logprint("- Done!")
    return 0

if __name__ == '__main__':
    root = Tk()
    root.title("Variant Counter")
    dfname = StringVar()
    idname = StringVar()
    ofname = StringVar()

    try:
        dfname.set(pickle.load(open( "pref.dat", "rb" )))
    except:
        pass

    frm = ttk.Frame(root, padding=10)
    frm.pack(side=LEFT)

    row1 = ttk.Frame(frm)
    ttk.Button(row1, width=25, text="Choose dictionary file:", command=openfile).pack(side=LEFT,padx=5)
    entry1 = ttk.Entry(row1,width=40,textvariable=dfname)
    row1.pack(side=TOP, padx=5, pady=5)
    entry1.pack(side=RIGHT, expand=YES, fill=X)
    entry1.xview_moveto(1)

    row2 = ttk.Frame(frm)
    ttk.Button(row2, width=25, text="Choose input directory:", command=opendir).pack(side=LEFT,padx=5)
    entry2 = ttk.Entry(row2,width=40,textvariable=idname)
    row2.pack(side=TOP, padx=5, pady=5)
    entry2.pack(side=RIGHT, expand=YES, fill=X)

    row3 = ttk.Frame(frm)
    ttk.Button(row3, width=25, text="Save output as:", command=savefile).pack(side=LEFT,padx=5)
    entry3 = ttk.Entry(row3, width = 40, textvariable=ofname)
    row3.pack(side=TOP, padx=5, pady=5)
    entry3.pack(side=RIGHT, expand=YES, fill=X)

    buttonrow = ttk.Frame(frm)
    ttk.Button(buttonrow, text="Run", command=run).pack(side=LEFT,padx=15)
    ttk.Button(buttonrow, text="Close", command=close).pack(side=RIGHT,padx=15)
    buttonrow.pack(side=BOTTOM,pady=30)
    
    log = Text(height=20,width=50)
    log.pack(side=RIGHT, padx=10,pady=10)

    root.protocol("WM_DELETE_WINDOW", close)

    root.mainloop()
