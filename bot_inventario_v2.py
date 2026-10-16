"""
Bot de Descarga de Inventarios IEEPO - Versión 2.0 Pro (Interfaz Gráfica Completa)
Lanzador principal que inicia la aplicación de escritorio GUI.
"""
import sys
import tkinter as tk
from gui_app import IEEPOBotGUI

def main():
    root = tk.Tk()
    app = IEEPOBotGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
