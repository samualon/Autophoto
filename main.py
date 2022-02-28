import win32com.client
import os


psApp = win32com.client.Dispatch("Photoshop.Application")

psApp.Open(r"C:\Users\Samuel\OneDrive - KU Leuven\Persoonlijk\Coding Projects\Python\Photoshop-automation\test.psd")

doc = psApp.Application.ActiveDocument

text_layer = doc.ArtLayers["custom"]
text_of_layer = text_layer.TextItem
text_of_layer.contents = "werkt!"
