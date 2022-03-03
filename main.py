import win32com.client
import os
from os import path
import psutil
import time

#Check if Ps is running
if("Photoshop.exe" in (i.name() for i in psutil.process_iter())):
    os.system("TASKKILL /F /IM Photoshop.exe")

#Check if result.png already exists
if(path.exists("result.png")):
    os.remove("result.png")

time.sleep(2)

#Locating template
print("Enter the location of the template (+ PSD name): ")
templateLocation = input()

#Dispatch
psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Open(templateLocation)
doc = psApp.Application.ActiveDocument

#Text and layer input
print("Enter the layer name which should be changed: ")
desiredLayer = input()

print("Enter the desired text: ")
desiredText = input()


time.sleep(2)

#Edit text
text_layer = doc.ArtLayers[desiredLayer]
text_of_layer = text_layer.TextItem
text_of_layer.contents = desiredText

#Export to png
options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
options.Format = 13   # PNG Format
options.PNG8 = False  # Sets it to PNG-24 bit

pngfile = r"C:\Users\Samuel\OneDrive - KU Leuven\Persoonlijk\Coding Projects\Python\Photoshop-automation\result.png"

doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)

#End check
print("Done!")
