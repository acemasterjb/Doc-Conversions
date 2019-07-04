from os import environ
from tkinter import filedialog
from mammoth import convert_to_html


"""
This is a script that optimizes the mammoth library's handling of 
docx styles to html tags

Using this script, Title 'paragraphs' in docx files will convert to *h1:fresh,
subtitle to *h2:fresh,  **v:stroke to a i.right element and further styling is
done in a separate style sheet

* :fresh meaning that each instance of the tag is unique, and not appended
with the previously converted one

** pending implementation
"""

# This is the file that the user selects; saves as a string
fileObj = filedialog.askopenfilename(initialdir = environ['USERPROFILE'],
                                     title = "Select file",
                                     filetypes = (("docx files","*.docx"),
                                                  ("doc files","*.doc"),
                                                  ("all files","*.*")))

# style maps, defined by Mammoth, map docx style to HTML ones
# essentially, they translate docx styles to HTML tags
style_map = """
            p[style-name='Title'] => h1:fresh
            p[style-name='Subtitle'] => h2:fresh
            """

# Converted file; saved as a object with a string formated with html tags and a message;
# omitts head and body tags
convFile = convert_to_html(fileObj, style_map=style_map)

# Reads output file, preformated with head and body tags
with open("output.html", "r") as htmlFile:
    htmlIn = htmlFile.read()

# Holds the string formatted with html tags
htmlOut = convFile.value

# Injects converted docx into body of output's *content*
html = htmlIn[:htmlIn.find("<body>") + 6] + htmlOut + htmlIn[htmlIn.find("<body>"):]

# Asks user *where* to save output file to; auto saves as a html file
out = filedialog.asksaveasfile(mode='w', defaultextension='.html')

# Writes html string into output file
out.write(html)
out.close()

print(convFile.messages)
