
import tkinter as tk
import tkinter.font
import tkinter.filedialog
import tkinter.messagebox

from tkinter import *
from pandastable import Table, TableModel, config
from tkinter import *
from pandastable import Table
#assuming parent is the frame in which you want to place the table
root=Tk()
parent=Frame(root)
pt = Table(parent)
# pt.show()
root.mainloop()
class TestApp(Frame):
        """Basic test frame for the table"""
        def __init__(self, parent=None):
            self.parent = parent
            Frame.__init__(self)
            self.main = self.master
            self.main.geometry('600x400+200+100')
            self.main.title('Table app')
            f = Frame(self.main)
            f.pack(fill=BOTH,expand=1)
            df = TableModel.getSampleData()
            self.table = pt = Table(f, dataframe=df,
                                    showtoolbar=True, showstatusbar=True)
            pt.show()
            #set some options
            options = {'colheadercolor':'green','floatprecision': 5}
            config.apply_options(options, pt)
            pt.show()
            return
app = TestApp()
#launch the app
app.mainloop()