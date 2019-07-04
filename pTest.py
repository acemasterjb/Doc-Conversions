from os import environ
from tkinter import filedialog
from pandas import read_excel
from pandas import DataFrame

"""
This script converts an xlsx table to a html table using the pandas library

Currently, this script converts to a simple html table with minimal styling

Optimization pending
"""

# This file obj is simply a string that holds the path for the selected Excel file
fileObj = filedialog.askopenfilename(initialdir=environ["USERPROFILE"],
                                     title="Select file",
                                     filetypes=(("xlsx files","*.xlsx"),
                                                ("xls files","*.xls"),
                                                ("all files","*.*")))

# This statement reads the file path and assigns to it a panda Excel object
tables = read_excel(fileObj)

# Convert the Excel object to html format, and save it as a file called tOut.html
html = tables.to_html('tOut.html')