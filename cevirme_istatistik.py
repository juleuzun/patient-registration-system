from PyQt6 import uic
import os

klasor = os.path.dirname(__file__)

ui_dosyasi = os.path.join(klasor, "istatistik.ui")
py_dosyasi = os.path.join(klasor, "istatistik.py")

with open(py_dosyasi, "w", encoding="utf-8") as f:
    uic.compileUi(ui_dosyasi, f)