Sub WorkingWithTextBoxes()

Dim PPTSlide As Slide
Dim TexFrm As TextFrame
Dim TexRng As TextRange

'Set the slide
Set PPTSlide = ActivePresentation.Slides(1)
Set TexFrm = PPTSlide.Shapes(1).TextFrame
Set TexRng = TexFrm.TextRange

'Does the Text frame have text?
Debug.Print TexFrm.HasText

'Change the orientation
TexFrm.Orientation = msoTextOrientationHorizontal

'Change the margin
TexFrm.MarginBottom = 40

'No Text Wrap
TexFrm.WordWrap = msoFalse

'Turn off AutoSize
TexFrm.AutoSize = ppAutoSizeNone

'How many characters
Debug.Print TexRng.Length

'Get the text
Debug.Print TexRng.Text

'Change the case
TexRng.ChangeCase (ppCaseUpper)

'Add a period
TexRng.AddPeriods

'Select certain characters
TexRng.Characters(Start:=1, Length:=2).Select

'Change the font
TexRng.Font.Name = "Roboto"

'Add some shadow
TexRng.Font.Shadow = msoTrue

'Find a word and delete it
TexRng.Words(Start:=1, Length:=1).Find("This").Delete

'Insert a slide number
TexRng.InsertSlideNumber

'Change the font Color
TexRng.Font.Color.ObjectThemeColor = msoThemeColorFollowedHyperlink

'Underline the text
TexRng.Font.Underline = True

End Sub
