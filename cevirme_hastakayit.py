from PyQt6 import uic
import os

klasor = os.path.dirname(__file__)

ui_dosyasi = os.path.join(klasor, "hastakayit.ui")
py_dosyasi = os.path.join(klasor, "hastakayit.py")

with open(py_dosyasi, "w", encoding="utf-8") as dosya:
    uic.compileUi(ui_dosyasi, dosya)

print("UI başarıyla çevrildi.")
    
    