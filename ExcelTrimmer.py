#!/usr/bin/env python
# coding: utf-8
# If pandas doesn't exist the requirements file will be read and pip will install dependencies
try:
    import pandas as panda
except:
    import sys
    import subprocess
    import pathlib
    path = str(pathlib.Path(__file__).parent.resolve())
    path += r"\requirements.txt"
    subprocess.check_call([sys.executable, '-m', 'pip','install','-r',path])
    import pandas as panda

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as filedialog
from datetime import datetime

def setLabelText(filepath):
    splitText = filepath.split("/")
    fileLabelText = splitText[-1]
    fileLabel.config(text = fileLabelText)

def setProccessedLabel(outputName):
    processedLabel.config(text="Has been proccessed and saved to: " + outputName)
    return

# Used to find file to trim
def handle_click(event):
    filepath = filedialog.askopenfilename(
        filetypes=[("CSV",".csv")]
    )
    if not filepath:
        return
    setLabelText(filepath)
    trim(filepath)

def trim(filepath):
    data = panda.read_csv(filepath)
    data.drop(["Internal ID","Document Number"],axis=1,inplace=True)
    groupedData = data.groupby("Item Description").agg({"Item Code": "first","Country of Manufacturer": "first","QTY": "sum","Uplift %":"first","VALUE PER ITEM (AUD)": "first"})
    groupedData.reset_index(inplace=True)

    def checkDescription(x):
        description = x["Item Description"]
        if "tissue" in description.lower():
            return "BOLLE SAFETY TISSUES"
        elif "cleaning spray" in description.lower():
            return "BOLLE SAFETY CLEANING SPRAY"
        else:
            return "BOLLE SAFETY EYEWEAR"

    groupedData["Description of Contents"] = groupedData.apply(checkDescription,axis=1)
    order = ["Item Code","Item Description", "Description of Contents","Country of Manufacturer","QTY","Uplift %","VALUE PER ITEM (AUD)"]
    groupedData = groupedData[order]
    
    path = filepath.split(r"/")
    path = path[:-1]

    now = datetime.now()
    #dd/mm H:M
    dtString =now.strftime("_%d%m_%H%M") 
    outputString = "OUTPUT" + dtString

    filepath = "/".join(path)
    filepath += "/" + outputString + ".xlsx"
    setProccessedLabel(outputString + ".xlsx")
    groupedData.to_excel(filepath,index=False,)

window = tk.Tk()
window.title("Simple Trimmer/Combiner")
window.eval('tk::PlaceWindow . center')
fileLabel = ttk.Label()
processedLabel = ttk.Label(
    #text="Has been processed to: "
)
pathButton = ttk.Button(
    text="Add & Trim CSV",
    width=70
)

window.bind("<Button-1>", handle_click)
fileLabel.pack()
processedLabel.pack()
pathButton.pack()

window.mainloop()

