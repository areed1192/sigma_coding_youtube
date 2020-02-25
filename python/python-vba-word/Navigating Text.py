import win32com.client as win32

# Grab the Word Application
WrdApp = win32.GetActiveObject("Word.Application")

# Grab the Document in the Applcation.
WrdDoc = WrdApp.Documents("MyDocument.docx")

'''
    The General Heirarachy is as follows:
    LEVEL 1: Sections
    LEVEL 2: Paragraphs
    LEVEL 3: Sentences
    LEVEL 4: Words
    LEVEL 5: Characters

    Additionally, sometimes to reach the next level down, means grabbing the range property
    then specifying the level you want.
'''

# Grab all the sections
WrdDocSecs = WrdDoc.Sections
print("There are {} SECTION(S) in the Document".format(WrdDocSecs.Count))

# Grab all the Paragraphs in that section, note I speicifed section 1, then the range of that section. From there, I specified the paragraphs collection.
SectIndex = 1
WrdDocParas = WrdDocSecs(SectIndex).Range.Paragraphs
print("There are {} PARGARPH(S) in SECTION {}".format(WrdDocParas.Count, SectIndex))

# Grab all the Sentences in Paragraph 1.
ParaIndex = 1
WrdDocSent = WrdDocParas(ParaIndex).Range.Sentences
print("There are {} SENTENCE(S) in PARAGRAPH {}".format(WrdDocSent.Count, ParaIndex))

# Grab all the Words in Senctence 1. NOTE THAT I DID NOT SPECIFY THE RANGE PROPERTY. THATS BECAUSE A SENTENCE IS A RANGE OBJECT.
SentIndex = 1
WrdDocWords = WrdDocSent(SentIndex).Words
print("There are {} WORD(S) in SENTENCE {}".format(WrdDocWords.Count, SentIndex))

# Grab all the characters in Senctence 1. AGAIN, NOTE I DID NOT SPECIFY THE RANGE OBJECT.
WordIndex = 1
WrdDocChara = WrdDocWords(WordIndex).Characters
print("There are {} CHARACTER(S) in WORD {}".format(WrdDocChara.Count, WordIndex))

# Define the Word
MyWord = WrdDocWords(WordIndex)

# Format the Word.
MyWord.Bold = True
MyWord.Italic = True
MyWord.Font.Name = "Arial Nova"
MyWord.Font.Size = 20