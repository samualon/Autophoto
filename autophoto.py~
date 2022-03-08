import win32com.client
import os
from os import path
import psutil
import time
import sys



#Arguments check
if(len(sys.argv) != 4):
    print("Fautly arguments! Correct syntax: autophoto.py 'doc_location' 'desired_layer' 'desired_text'.")
    print(len(sys.argv))
    for i in sys.argv:
        print(i)
    exit()



#Pre startup checks
if("Photoshop.exe" in (i.name() for i in psutil.process_iter())):
    os.system("TASKKILL /F /IM Photoshop.exe")

if(path.exists("result.png")):
    os.remove("result.png")

time.sleep(2)



#Argument declaration
templateLocation = str(sys.argv[1])
desiredLayer = str(sys.argv[2])
desiredText = str(sys.argv[3])



#Dispatch
psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Open(templateLocation)
doc = psApp.Application.ActiveDocument



#Editing
text_layer = doc.ArtLayers[desiredLayer]
text_of_layer = text_layer.TextItem
text_of_layer.contents = desiredText



#Exporting
options = win32com.client.Dispatch('5516.ExportOptionsSaveForWeb')
options.Format = 13   # PNG Format
options.PNG8 = False  # Sets it to PNG-24 bit

pngfile = r"C:\Users\Samuel\OneDrive - KU Leuven\Persoonlijk\Coding Projects\Python\Photoshop-automation\result.png"

doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)



#End check
print("Done!")
