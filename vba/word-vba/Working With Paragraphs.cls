Sub EditParagraph()

Dim WrdPara As Paragraph
Dim ActDocm As Document

'Get the active document & the paragraph we want to work with
Set ActDocm = ActiveDocument
Set WrdPara = ActDocm.Paragraphs(1)

'Set the Alignment
WrdPara.Alignment = wdAlignParagraphJustifyMed

'Add a border around our Paragraph
WrdPara.Borders.Shadow = True
WrdPara.Borders.OutsideLineStyle = wdLineStyleDashDot
WrdPara.Borders.OutsideColor = wdColorAqua

'Add a drop cap
WrdPara.DropCap.Enable

'Remove a drop cap
WrdPara.DropCap.Clear

'Add line spacing
WrdPara.LineSpacing = 20

'Word defined Line Spacing
WrdPara.LineSpacingRule = wdLineSpaceDouble

'Count the number of Words in our Paragraph
Debug.Print WrdPara.Range.Words.Count

'Select the first word
WrdPara.Range.Words.First.Select
WrdPara.Range.Words.Last.Select

'Bold the Paragraph
WrdPara.Range.Bold = True

'Select the first letter in a paragraph
WrdPara.Range.Characters.First.Select

'Change all the text to uppercase
WrdPara.Range.Case = wdUpperCase

'Capitalize Each Word
WrdPara.Range.Case = wdTitleWord

End Sub
