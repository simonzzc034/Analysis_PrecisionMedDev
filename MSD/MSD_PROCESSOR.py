import pandas as pd
import time
import tkinter as tk 
import tkinter.ttk as ttk
from tkinter import messagebox
from tkinter.filedialog import askopenfile

root = tk.Tk()

root.title('MSD Data Processor')
root.geometry("650x650")
root.config(background = "white")

upload_button = tk.Button(root,
    text="Choose Input File",
    command=lambda:UploadAction(),
    width=40,
    height=5
    ).pack()

info = tk.Label(text = "This tool formats MSD Workbench output into a graph (Prism) ready pivot table.\n\nIt takes a MSD output file as input, \n\nand write the pivot tables to a file in the same folder as the input"\
        "\n\n Click the button above to choose your input file",
             foreground='blue',
             background='white',
             height=20,
             relief=tk.RAISED,
             borderwidth=5,
             font=("Arial", 12, "bold")
            )
info.configure(anchor="center")
info.pack()

def UploadAction():
    file_path = askopenfile(mode='r', filetypes=[('MSD Output Files','*xlsx')])
    output_file = file_path.name.replace("xlsx","pivot.table.xlsx")

    raw_data = pd.read_excel(file_path.name,skiprows=1)
    data_list = []
    for i in range(0,int((raw_data.shape[1]+1)/7)):
        temp = raw_data.iloc[:,i*7:(i+1)*7-1]
        temp.columns = ['Sample','Calc.Concentration','Assay','Detection Range','Dilution', 'Well']
        data_list.append(temp)

    x = pd.concat(data_list)
    x = x.loc[~x['Sample'].str.startswith('S0'),:]
    x[['Sample','Barcode','Treatment','Dosage','Cell_Count','Replicate','Time']] = (x['Sample'].str.split(",", expand=True)).iloc[:,0:7]
    x = x[['Sample','Assay','Treatment','Time','Replicate','Calc.Concentration','Dosage','Barcode','Detection Range','Dilution','Cell_Count']]
    x_table = x.groupby(['Sample', 'Assay','Treatment','Replicate','Time'])['Calc.Concentration'].aggregate(['mean']).unstack().unstack()
    #x_table = x_table.reindex(index = x_table.index.reindex([' Not Treated',' IL-2',' LPS'], level="Treatment")[0], columns = x_table.columns.reindex([' 2 hours',' 4 hours',' 6 hours',' 16 hours',' 20 hours', ' 24 hours'],level='Time')[0])
    x_table = x_table.reindex(columns = x_table.columns.reindex([' 2 hours',' 4 hours',' 6 hours',' 16 hours',' 20 hours', ' 24 hours'],level='Time')[0])
    x_table.columns.names = ['Calc.Concentration', 'Time', 'Replicate']

    x_table.to_excel(output_file)

    time.sleep(1)
    messagebox.showinfo('Done!','Data finished processing, file written to ' + output_file)

root.mainloop()
