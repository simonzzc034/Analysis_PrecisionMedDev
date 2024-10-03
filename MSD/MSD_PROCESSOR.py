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

info = tk.Label(text = "This tool formats MSD Workbench output into a graph (Prism) ready format.\n\nIt takes a MSD 'Raw Plate Reads' as input, \n\nand write the output to a file in the same folder as the input"\
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
    output_file = file_path.name.replace("xlsx","processedForPrism.xlsx")

    raw_data = pd.read_excel(file_path.name,sheet_name = 'Raw Plate Reads',skiprows=1)
    data_list = []
    for i in range(0,int((raw_data.shape[1]+1)/7)):
        temp = raw_data.iloc[:,i*7:(i+1)*7-1]
        temp.columns = ['Sample','Assay','Well','Detection Range','Dilution','Calc.Concentration']
        
        data_list.append(temp)

    x = pd.concat(data_list)
    x = x.loc[x['Sample'].notna(),:]
    x = x.loc[~x['Sample'].str.startswith('S0'),:]
    x[['Sample','Experiment','Drug','Concentration','Stim','REP']] = (x['Sample'].str.split(",", expand=True)).iloc[:,0:6]
    x = x[['Sample','Experiment','Drug','Concentration','Stim','REP','Assay','Well','Detection Range','Dilution','Calc.Concentration']]

    cutoffs = dict()
    cutoffs[' DK210E'] = dict()
    cutoffs[' DK710'] = dict()
    cutoffs[' DK1210'] = dict()
    cutoffs[' DKalpha10'] = dict()

    cutoffs[' DK210E']['low'] = 3000
    cutoffs[' DK210E']['high'] = 16000
    cutoffs[' DK710']['low'] = 5000
    cutoffs[' DK710']['high'] = 16000
    cutoffs[' DK1210']['low'] = 7000
    cutoffs[' DK1210']['high'] = 20000
    cutoffs[' DKalpha10']['low'] = 2000
    cutoffs[' DKalpha10']['high'] = 4000

    high = dict()
    medium = dict()
    low = dict()
    response = dict()
    
    for drug in x['Drug'].unique(): 
        high[drug] = []
        medium[drug] = []
        low[drug] = []
        
        for sample in x.loc[x['Drug']==drug, 'Sample'].unique():
            gamma = 0.25 * x.query('Assay == "IFN-γ" and Drug==@drug and Concentration==" 20 ng/mL" and Sample == @sample ')['Calc.Concentration'].mean()
            gamma += 0.75 * x.query('Assay == "IFN-γ" and Drug==@drug and Concentration==" 200 ng/mL" and Sample == @sample ')['Calc.Concentration'].mean()

            if gamma < cutoffs[drug]['low']:
                low[drug].append(sample)
            elif gamma > cutoffs[drug]['high']:
                high[drug].append(sample)
            else:
                medium[drug].append(sample)
        response[drug] = pd.DataFrame({'High': pd.Series(high[drug], dtype='object'), 'Medium': pd.Series(medium[drug],dtype='object'), 'low': pd.Series(low[drug], dtype='object')})   

    with pd.ExcelWriter(output_file) as output_xls:
    
        for assay in x['Assay'].unique():
            for drug in x['Drug'].unique():
                if len(high[drug]) > 0:
                    temp_data_high = x.loc[(x['Assay']==assay) & (x['Drug']==drug) & (x['Sample'].isin(high[drug])),:]
                    temp_table_high = temp_data_high.groupby(['Concentration', 'Sample'])['Calc.Concentration'].aggregate(['mean']).unstack()
                    temp_table_high.columns.names = ['High', 'Sample']
                    temp_table_high.index.names = ['']
                    temp_table_high.to_excel(output_xls, sheet_name = assay + ' ' + drug)
            
                if len(medium[drug])>0:
                    temp_data_medium = x.loc[(x['Assay']==assay) & (x['Drug']==drug) & (x['Sample'].isin(medium[drug])),:]
                    temp_table_medium = temp_data_medium.groupby(['Concentration', 'Sample'])['Calc.Concentration'].aggregate(['mean']).unstack()
                    temp_table_medium.columns.names = ['Medium', 'Sample']
                    temp_table_medium.index.names = ['']
                    temp_table_medium.to_excel(output_xls, sheet_name = assay + ' ' + drug, startrow=7)
            
                if len(low[drug])>0:
                    temp_data_low = x.loc[(x['Assay']==assay) & (x['Drug']==drug) & (x['Sample'].isin(low[drug])),:]
                    temp_table_low = temp_data_low.groupby(['Concentration', 'Sample'])['Calc.Concentration'].aggregate(['mean']).unstack()
                    temp_table_low.columns.names = ['Low', 'Sample']
                    temp_table_low.index.names = ['']
                    temp_table_low.to_excel(output_xls, sheet_name = assay + ' ' + drug, startrow=14)
            
                response[drug].to_excel(output_xls, sheet_name = assay + ' ' + drug, startrow=22)


    time.sleep(1)
    messagebox.showinfo('Done!','Data finished processing, file written to ' + output_file)

root.mainloop()
