
import win32com.client as win32

# CREATE DOCUMENT - LATE BINDING.

# Create an instance of the Word App using Early Binding
WordApp = win32.gencache.EnsureDispatch('Word.Application')

# make the app visible
WordApp.Visible = True

# add a document.
WrdDoc = WordApp.Documents.Add()


# create a Word Document com object using Early Binding.
WrdDoc = win32.gencache.EnsureDispatch(WordApp.Documents(1))

# because I used early binding, I can see all of the methods.
help(WrdDoc)

# I can also get the attributes of the objects within it.
help(WrdDoc.__getattr__('Paragraphs'))



from bs4 import BeautifulSoup
import requests

html_code = requests.get('https://en.wikipedia.org/wiki/Star_Wars').content

# dump the code in a BeautifulSoup parser
soup = BeautifulSoup(html_code, 'html.parser')

# let's look for all the `images`
images = soup.find_all('img')

# Get the number of rows we will need
numrows = len(images)

# Add a table to the document and set the style
WrdTbl = WrdDoc.Tables.Add(WrdDoc.Paragraphs(1).Range, numrows, 2)
WrdTbl.Style = "Grid Table 1 Light - Accent 1"

# Populate the header rows
WrdTbl.Cell(1, 1).Range.Text = "Link Number"
WrdTbl.Cell(1, 2).Range.Text = "Link "

# loop through each of the images, add a paragrph and put the link the document.
for index, img in enumerate(images):    
    
    # populate each cell in the table with the link and link ID
    WrdTbl.Cell(index + 2, 1).Range.Text = "Link: " + str(index)
    WrdTbl.Cell(index + 2, 2).Range.Text = img['src'].replace(r'//', 'https://')

