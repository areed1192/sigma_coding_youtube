import pprint
import win32com.client as win32
from win32com.client import constants as win_const

# Grab the Active Instance of Adobe.
try:
    adobe_app = win32.GetActiveObject("Illustrator.Application")
except:
    adobe_app = win32.gencache.EnsureDispatch("Illustrator.Application")

# Define the Document we will be working with.
try:
    adobe_doc = adobe_app.ActiveDocument
except:
    adobe_doc = adobe_app.Documents.Add(
        win_const.aiDocumentCMYKColor,
        Width=300,
        Height=300
    )

# Grab the Documents Collection.
adobe_docs = adobe_app.Documents

# Grab the Document Count.
print(
    "The number of Documents are {doc_count}".format(
        doc_count=adobe_docs.Count
    )
)

# Add a new Document.
adobe_doc_new = adobe_docs.Add(
    win_const.aiDocumentCMYKColor,
    Width=300,
    Height=300
)

# Print the New Document Name.
print(
    "The new document Name is {new_doc_name}".format(
        new_doc_name=adobe_doc_new.Name
    )
)

# Close that Document.
adobe_doc_new.Close()

# Print the File Name.
print("File Name: {name}".format(name=adobe_doc.Name))

# Print the File Path.
print("File Path: {path}".format(path=adobe_doc.Path))

# Grab the Active Layer.
active_layer = adobe_doc.ActiveLayer

# Print the File Path.
print("Active Layer Name: {Name}".format(Name=active_layer.Name))

# Grab the Page Origin.
print(adobe_doc.PageOrigin)

# Grab the Page Hieght.
print(adobe_doc.Height)

# Grab the Page Width.
print(adobe_doc.Width)

# Is the Document Saved?
print(adobe_doc.Saved)

# What is the output resolution.
print(adobe_doc.OutputResolution)

# Grab the Document Color Space Enumeration.
print(adobe_doc.DocumentColorSpace)

# What is the Default Stroke Width?
print(adobe_doc.DefaultStrokeWidth)

# What is the Cyan Value for the Default Fill Color.
print(adobe_doc.DefaultFillColor.Cyan)

# Grab the Symbols.
doc_symbols = adobe_doc.Symbols
print(doc_symbols.Count)

# Grab the Symbols Items Collection.
doc_symbols_items = adobe_doc.SymbolItems
print(doc_symbols_items.Count)

# Grab the Variables Collection.
doc_vars = adobe_doc.Variables
print(doc_vars.Count)

# Grab the TextFrames Collection.
doc_text_frames = adobe_doc.TextFrames
print(doc_text_frames.Count)

# Grab the Tags Collection.
doc_tags = adobe_doc.Tags
print(doc_tags.Count)
