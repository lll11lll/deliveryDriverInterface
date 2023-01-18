# Nick Greiner
# 01/17/2023
# Setting Frame Class for the driver interface, allows the user to open excel,
# and quit the program

import tkinter as tk
from tkinter import ttk
import subprocess


class SettingFrame(ttk.LabelFrame):
    def __init__(self, mainFrame):
        super().__init__(mainFrame)

        # Giving the frame a title
        self['text'] = 'Open Excel'

        options = {'padx': 10, 'pady': 5}

        self.openExcelSheetButton = tk.Button(
            self, text='Open Excel', width=50, bg='blue')

        def openFile():
            subprocess.Popen('myDriverData2023.xlsx', shell=True)

        self.openExcelSheetButton.configure(command=openFile)
        self.openExcelSheetButton.grid(
            row=0, column=0, rowspan=2, columnspan=2, ** options)

        self.grid(row=2, column=0, columnspan=3,
                  sticky="ew", **options)
