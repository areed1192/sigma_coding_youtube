Sub UsingTheFindObject_Simple()

''
'' SIMPLE EXAMPLE OF USING THE FIND OBJECT TO FIND THE INSTANCE OF THE WORD.
''

'Declare Variables.
Dim wrdFind As Find
Dim wrdRng As Range
Dim wrdDoc As Document

'Grab the ActiveDocument.
Set wrdDoc = Application.ActiveDocument

'Define the Content in the document
Set wrdRng = wrdDoc.Content

'Define the Find Object based on the Range.
Set wrdFind = wrdRng.Find

'Define the paramaters of the Search.
With wrdFind
    
    'Let's start simple and Find the text "Jeff Bezos"
    .Text = "Jeff Bezos"
    
    'Conduct the search, and store the results of that search in a variable. Returns TRUE for a match and FALSE for no match.
    searchResult = .Execute
    
End With

'If we found it, display a message and then set all the results to BOLD.
If searchResult = True Then

    'Display a message.
    Debug.Print "Found the word Jeff Bezos, now formatting."
    
    'Change the font to bold.
    wrdRng.Bold = True

End If

End Sub

Sub UsingTheFindObject_Medium()

''
'' MEDIUM EXAMPLE OF USING THE FIND OBJECT TO FIND THE INSTANCE OF THE
'' WORD AND FORMAT THAT SPECIFIC INSTANCE BASED ON IT'S EXISTING OBJECT.
''

'Declare Variables.
Dim wrdFind As Find
Dim wrdRng As Range
Dim wrdDoc As Document

'Grab the ActiveDocument.
Set wrdDoc = Application.ActiveDocument

'Define the Content in the document
Set wrdRng = wrdDoc.Content

'Define the Find Object based on the Range.
Set wrdFind = wrdRng.Find

'Define the paramaters of the Search.
With wrdFind
    
    'Let's search for the text "Amazon"
    .Text = "Amazon"
    
    'Must match the casing.
    .MatchCase = True
    
    'Also it need to match the whole world.
    .MatchWholeWord = True
    
    'Also I want any formatting search rules to be followed.
    .Format = True
    
End With

'In this case I want to find ALL the instances of the word AMAZON, so let's use a while loop that will keep calling the Execute method
'until it returns FALSE.
Do While wrdFind.Execute = True

    'Display a message.
    Debug.Print "Found the word Amazon, now formatting."
    
    'Change the font to bold.
    wrdRng.Bold = True
    
    'Change the color to Blue.
    wrdRng.Font.ColorIndex = wdBlue
    
    'We could also format the paragraph that the range is in.
    wrdRng.Paragraphs.Alignment = wdAlignParagraphJustify
    
    'However, if my range that is found has a hyperlink in it. I want it colored Green.
    If wrdRng.Hyperlinks.Count > 0 Then
    
        'Change the font to green.
        wrdRng.Font.ColorIndex = wdGreen
        
    End If

    'If it's part of a bullet list then we need to format it red.
    If wrdRng.ListFormat.ListType = wdListBullet Then

        'Change the font to green.
        wrdRng.Font.ColorIndex = wdRed

    End If
    
Loop

End Sub

Sub UsingTheFindObject_Complex()

''
'' COMPLEX EXAMPLE: DOING MULTIPLE SEARCHES AND DEFINING SEPERATE
'' FORMATTING CONDITIONS FOR EACH SEARCH.
''

'Declare Variables.
Dim wrdFind As Find
Dim wrdRng As Range
Dim wrdDoc As Document

'Define some arrays to store our info.
Dim wrdSearchList As Variant
Dim wrdSearchFromat As Variant

'Grab the ActiveDocument.
Set wrdDoc = Application.ActiveDocument

'Create an array of strings that we will search for.
wrdSearchList = Array(">%", "$*")

'Define a list of Colors to highlight
wrdSearchFormat = Array(wdColorOrange, wdColorBrightGreen)

'Step 1: Loop through the array of symbols we want to search.
For i = LBound(wrdSearchList) To UBound(wrdSearchList)

    'The Tricky part is you have to reset the variables on every loop or it will only do the first search.
    'Why you might ask? Because if there is a found result the wrdRng will be redefined to include only the matches.
    
    'Define the Content in the document
    Set wrdRng = wrdDoc.Content
    
    'Define the Find Object based on the Range.
    Set wrdFind = wrdRng.Find

    'Define the paramaters of the Search.
    With wrdFind
        
        'Search for the element we are on in the loop.
        .Text = wrdSearchList(i)
        
        'We are usign wildcards so we need to set that property to True.
        .MatchWildcards = True

    End With
    
    'Keep going as long as there is a match.
    Do While wrdFind.Execute = True
    
        'Display a message.
        Debug.Print "Found the word " + wrdSearchList(i) + " now formatting."
        
        'Change the font to bold.
        wrdRng.Bold = True
        
        'Change the color to match the one on the loop.
        wrdRng.Font.Color = wrdSearchFormat(i)
        
    Loop
    
Next

End Sub