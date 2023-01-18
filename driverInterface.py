# Nick Greiner
# 01/16/2023
# I want to recreate the driver interface I created with tkinter however I want
# to have a objected oriented approach using classes

import tkinter as tk
from tkinter import ttk
import sv_ttk
from mainFrameClass import *


class DriverInterface(tk.Tk):
    def __init__(self):
        super().__init__()

        # Configuring the main interface
        self.title('Delivery Driver Interface')


if __name__ == '__main__':
    myInterface = DriverInterface()

    sv_ttk.set_theme('dark')
    MainFrame(myInterface)

    myInterface.mainloop()
