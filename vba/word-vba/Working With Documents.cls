Sub WorkingWithADocument()
Dim i As Long

'Count the lines
Debug.Print "The number of lines are: " + CStr(ActiveDocument.ComputeStatistics(Statistic:=wdStatisticLines))

'Count the pages
Debug.Print "The number of pages are: " + CStr(ActiveDocument.ComputeStatistics(Statistic:=wdStatisticPages))

'Count the Words
Debug.Print "The number of Words are: " + CStr(ActiveDocument.ComputeStatistics(Statistic:=wdStatisticWords))

'Count the Paragraphs
Debug.Print "The number of paragraphs are: " + CStr(ActiveDocument.ComputeStatistics(Statistic:=wdStatisticParagraphs))

'Count the characters, inlcuding spaces
Debug.Print "The number of characters are: " + CStr(ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharactersWithSpaces))

'Count the characters, not including spaces
Debug.Print "The number of characters are: " + CStr(ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharacters))

'Count the number of Grammatical Errors.
Debug.Print "The number of grammatical errors are: " + CStr(ActiveDocument.GrammaticalErrors.Count)

'Work with the errors in the document.
For i = 1 To ActiveDocument.GrammaticalErrors.Count
    Debug.Print ActiveDocument.GrammaticalErrors.Item(i).Text
    Debug.Print ActiveDocument.GrammaticalErrors.Item(i).SpellingErrors(1).Text
    Debug.Print ActiveDocument.GrammaticalErrors.Item(i).GrammaticalErrors(1).Text
Next

'Change The font for the entire document
ActiveDocument.Content.Font.Name = "Roboto"

'Remove the bullets and numbers.
ActiveDocument.RemoveNumbers

'Check if it has a password
Debug.Print ActiveDocument.HasPassword

'Check if it has a Macro
Debug.Print ActiveDocument.HasVBProject

'Get the path.
Debug.Print ActiveDocument.Path

'Turn off Grammar Checking
Options.CheckGrammarAsYouType = False
ActiveDocument.ShowGrammaticalErrors = False

'First let's make sure we display the background, by default we don't see it.
ActiveDocument.ActiveWindow.View.DisplayBackgrounds = True

'Set the background color of a document.
With ActiveDocument.Background.Fill
    .Visible = True
    .ForeColor.RGB = RGB(43, 87, 154)
End With

'Display the rulers
Application.ActiveWindow.ActivePane.DisplayRulers = True

'Display the document map
Application.ActiveWindow.DocumentMap = True

'Display the Gridlines
Application.Options.DisplayGridLines = True

'Set the zoom
With Application.ActiveWindow.ActivePane.View.Zoom
    .PageColumns = 3
    .PageRows = 2
End With

'Set it back to one page width view
Application.ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitFullPage

End Sub
