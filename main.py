import win32com.client
import os
import psutil

#Check if Ps is running
if("Photoshop.exe" in (i.name() for i in psutil.process_iter())):
    os.system("TASKKILL /F /IM Photoshop.exe")



#Dispatch
psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Open(r"C:\Users\Samuel\OneDrive - KU Leuven\Persoonlijk\Coding Projects\Python\Photoshop-automation\test.psd")
doc = psApp.Application.ActiveDocument

#Edit text
text_layer = doc.ArtLayers["custom"]
text_of_layer = text_layer.TextItem
text_of_layer.contents = "werkt!"

#Export to png
options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
options.Format = 13   # PNG Format
options.PNG8 = False  # Sets it to PNG-24 bit

pngfile = r"C:\Users\Samuel\OneDrive - KU Leuven\Persoonlijk\Coding Projects\Python\Photoshop-automation\result.png"

doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)

#End check
print("Done!")
