try:
    import pandas as panda
except:
    import sys
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip','install','-r','requirements.txt'])
    import pandas as panda

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as filedialog

def handle_click(event):
    filepath = filedialog.askopenfilename(
        filetypes=[("CSV",".csv")]
    )
    if not filepath:
        return
    trim(filepath)

def trim(filepath : str):
    sheet = panda.read_csv(filepath)
    sheet.drop(sheet.tail(1).index,inplace=True)

    myDict = {"Item Code" : [], "Item Description": [], "Country of Manufacturer" : [], "QTY": [], "VALUE PER ITEM (AUD)" : []}
    totalDict = {}
    otherDict = {"Item Code" : [], "Country of Manufacturer" : [],"VALUE PER ITEM (AUD)" : [],"Item Description": []}

    for index, row in sheet.iterrows():
        description = row["Item Description"]
        qty = row["QTY"]

        if description not in totalDict:
            totalDict[description] = int(qty)
            otherDict["Item Code"].append(row["Item Code"])
            otherDict["Item Description"].append(description)
            otherDict["Country of Manufacturer"].append(row["Country of Manufacturer"])
            otherDict["VALUE PER ITEM (AUD)"].append(row["VALUE PER ITEM (AUD)"])
        else:
            totalDict[description] += int(qty)

    count = 0
    for value in totalDict:
        description = otherDict["Item Description"][count]
        myDict["Item Code"].append(otherDict["Item Code"][count])
        myDict["Item Description"].append(description)
        myDict["Country of Manufacturer"].append(otherDict["Country of Manufacturer"][count])
        myDict["QTY"].append(totalDict[description])
        myDict["VALUE PER ITEM (AUD)"].append(otherDict["VALUE PER ITEM (AUD)"][count])
        count+=1

    df = panda.DataFrame(data=myDict)
    path = filepath.split(r"/")
    path = path[:-1]

    filepath = "/".join(path)
    filepath += "/output.xlsx"
    df.to_excel(filepath)
    x = 1


window = tk.Tk()
pathButton = ttk.Button(
    text="Add & Trim CSV",
    width=15
)

window.bind("<Button-1>", handle_click)
pathButton.pack()

window.mainloop()