Sub LoopThroughEachWordInDoc()

'Declare our Range Object
Dim WrdRng As Range

'Loop through each word in the document.
For Each WrdRng In ActiveDocument.Words
    
    'Trim the trailing space. This will take "Font " & make it "Font"
    WrdText = RTrim(wd.Text)
    
    'If the color is Black & the word is Font, then change it to bold & green
    If wd.Font.Color = vbBlack And WrdText = "Font" Then
       wd.Font.Bold = True
       wd.Font.Color = vbGreen
    End If
    
Next

End Sub
