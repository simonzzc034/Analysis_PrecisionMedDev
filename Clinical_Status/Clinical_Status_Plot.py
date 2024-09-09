import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.patches as patches
import time
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.filedialog import askopenfile

root = tk.Tk()

root.title('Clinical Patients Status Plotter')
root.geometry("650x650")
root.config(background = "white")

combo_title = tk.Label(root, text="Select a plot type:")
combo_title.pack()

types = ["External", "Internal"]
combobox = ttk.Combobox(root, values=types)
combobox.pack()

info = tk.Label(text = "This tool takes an Excel file containing the clinical trial patients status\n\nand makes a swimmer plot pdf file in the same folder as the input"\
        "\n\n Use the dropdown above to pick a plot type first and click the button below to choose your input file",
             foreground='blue',
             background='white',
             height=20,
             relief=tk.RAISED,
             borderwidth=5,
             font=("Arial", 12, "bold")
            )
info.configure(anchor="center")
info.pack()



upload_button = tk.Button(root,
    text="Choose Input File",
    command=lambda:UploadAction(),
    width=40,
    height=5
    ).pack()



def UploadAction():
    file_path = askopenfile(mode='r', filetypes=[('Patient Status File','*xlsx')])
    plot_type = combobox.get()
    if plot_type == "External":
        output_file = "swimmer.plot.pdf"
    else:
        output_file = "swimmer.plot.withID.pdf"

    patient_data = pd.read_excel(file_path.name)
    patient_data['Date for Treatment Span'] = patient_data['End of Treatment Date'].fillna(pd.to_datetime('today').normalize())
    patient_data['Weeks in Study(2)'] = (patient_data['Date for Treatment Span'] - patient_data['Start Date'])/pd.Timedelta(weeks=1)
    patient_data['Weeks in Study(2)'] = round(patient_data['Weeks in Study(2)'])
    patient_data = patient_data.sort_values('Weeks in Study(2)', ascending=False)
    patient_data.index = range(patient_data.shape[0])
    patient_data['Cohort Number'] = patient_data['Cohort Number'].replace('3 (8 mg) TIW', '3 (8 mg)')

    dosages = ['1 (2 mg)','2 (4 mg)','2.5 (6 mg)','3 (8 mg)','3.5 (10 mg)','4 (16 mg)']
    dosages_color = {
        '1 (2 mg)': 'orange',
        '2 (4 mg)': 'blue',
        '2.5 (6 mg)': 'green',
        '3 (8 mg)': 'grey',
        '3.5 (10 mg)': 'cyan',
        '4 (16 mg)': 'red'
    }

    status = ['Partial Response','Stable Disease','Pseudo Progression','Clinical Progression','Progressive Disease','On Study']
    status_marker = {
        'Partial Response': "P",
        'Stable Disease': "d",
        'Pseudo Progression': "*",
        'Clinical Progression': "o",
        'Progressive Disease': "D",
        'On Study': ">"
    }
    status_color = {
        'Partial Response': "black",
        'Stable Disease': "blue",
        'Pseudo Progression': "purple",
        'Clinical Progression': "orange",
        'Progressive Disease': "red",
        'On Study': "green"
    }
    status_shorts = {
        'PR': "Partial Response",
        'SD': "Stable Disease",
        'PD': "Progressive Disease",
        'Pseudo progression': "Pseudo Progression",
        'CP': 'Clinical Progression'
    }

    sns.set_style("white")
    fig, ax = plt.subplots()
    ax.set_aspect(0.5)
    
    if plot_type == "External":
        for i, patient in enumerate(patient_data['Cancer Type']):
            if patient_data.loc[i,'Weeks in Study(2)'] > 0:
                ax.text(-1, 2*i, patient, ha='right', va='center', fontsize=6)
        ax.set_ylabel("Patient Cancer Type", fontsize=8)
    else:
        for i, patient in enumerate(patient_data['Patient Number'] + ' (' + patient_data['Cancer Type'] + ')'):
            if patient_data.loc[i,'Weeks in Study(2)'] > 0:
                ax.text(-1, 2*i, patient, ha='right', va='center', fontsize=6)
    
    ax.set_title('Clinical Trial Patient Status', fontsize = 8)
    ax.set_yticklabels([])
    ax.yaxis.grid(False)
    ax.spines['left'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.tick_params(axis='x', labelsize=6, bottom=True, length=3)
    ax.set_xlim(0,patient_data.shape[1]-7)
    ax.set_xlabel("Weeks In Study", fontsize=8)
    ax.yaxis.set_label_coords(-0.1,0.5)
    
    ax.text(23,56, 'Dose level:', fontsize=8)
    
    for i in range(6):
        ax.text(23,52-2*i,dosages[i],fontsize=8)
        rect = patches.Rectangle((32, 52-2*i), 2, 1, linewidth=0.5, edgecolor='black', facecolor=dosages_color[dosages[i]], alpha=0.5)
        ax.add_patch(rect)
    
    ax.text(23,38, 'Status:', fontsize=8)
    
    for i in range(5):
        ax.text(23,34-2*i,status[i], fontsize=8)
        ax.scatter(35,34.5-2*i, marker=status_marker[status[i]], color=status_color[status[i]], s=15)
    
    ax.text(23, 24, status[5], fontsize=8)
    ax.arrow(34, 24, 1, 0, head_width=1, head_length=1, fc='g', ec='g' )
    
    for i in range(patient_data.shape[0]):
        if patient_data.loc[i,'Weeks in Study(2)'] > 0:
            anchor = patient_data.loc[i,'Weeks in Study(2)']
            rect = patches.Rectangle((0,i*2-0.8),anchor,1.5,linewidth=1, facecolor=dosages_color[patient_data['Cohort Number'][i]],edgecolor="black",alpha=0.3)
            ax.add_patch(rect)
    
            if pd.isnull(patient_data.loc[i, 'End of Treatment Date']):
                #ax.scatter(anchor-0.4, 2*i, marker='>', color='green', s=8)
                ax.arrow(anchor, 2*i, 1, 0, head_width=1, head_length=1, fc='g', ec='g' )
            for j in range(7,patient_data.shape[1]-2):
                if (not pd.isna(patient_data.iloc[i,j])) and (patient_data.iloc[i,j] in status_shorts):
                    ax.scatter(j-6.6,2*i,marker=status_marker[status_shorts[patient_data.iloc[i,j]]], color=status_color[status_shorts[patient_data.iloc[i,j]]], s=8)
    fig.savefig(output_file)

    time.sleep(1)
    messagebox.showinfo('Done!','Plot is made, see file ' + output_file)

root.mainloop()
