import win32com.client
import os


psApp = win32com.client.Dispatch("Photoshop.Application")

psApp.Open(r"C:\Users\Samuel\OneDrive - KU Leuven\Persoonlijk\Coding Projects\Python\Photoshop automation\test.psd")

doc = psApp.Application.ActiveDocument

layer_ = doc.ArtLayers["Facts"]
text_of_layer = layer_facts.TextItem
text_of_layer.contents = "This is an example of a new text."
