# Nick Greiner
# 01/16/2023
# This is the main frame for the driver interface window

import tkinter as tk
from tkinter import ttk
from dataEntryFrameClass import DataEntryFrame
from buttonFrameClass import ButtonFrame
from settingFrameClass import SettingFrame


class MainFrame(ttk.Frame):
    def __init__(self, DriverInterface):
        super().__init__(DriverInterface)
        options = {'padx': 10, 'pady': 5}

        self.dataEntryFrame = DataEntryFrame(self)
        self.buttonFrame = ButtonFrame(self)
        self.settingFrame = SettingFrame(self)

        self.quitButton = tk.Button(
            self, text='Quit', width=50, bg='red', command=DriverInterface.destroy, **options)

        # def destroyWindow():
        # DriverInterface.destroy()
        self.quitButton.grid(row=3, column=0, sticky='ew')
        self.grid(row=0, column=0)
