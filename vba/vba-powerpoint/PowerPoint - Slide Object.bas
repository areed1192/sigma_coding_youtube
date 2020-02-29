Option Explicit

Sub WorkWithSlides()

Dim PPTSlide As Slide
Dim PPTShape As Shape

Set PPTSlide = ActivePresentation.Slides(1)

'Work with the slide header and footer.
PPTSlide.HeadersFooters.DateAndTime.Visible = True
PPTSlide.HeadersFooters.SlideNumber.Visible = True
PPTSlide.HeadersFooters.Footer.Visible = True
PPTSlide.HeadersFooters.Footer.Text = "Hi there"

'Add a slide comment
PPTSlide.Comments.Add Top:=100, Left:=100, Author:="Alex", AuthorInitials:="AR", Text:="Looks Good"

'Print the slide name, depends on the slide position.
Debug.Print PPTSlide.Name

'Print some more details about the slide.
Debug.Print PPTSlide.SlideNumber
Debug.Print PPTSlide.SlideIndex
Debug.Print PPTSlide.SlideID

'Add a tag to the slide that the user can't see
PPTSlide.Tags.Add Name:="MyTag", Value:="Finance Slides"

'Read the tag, this isn't intiutive but we can do it.
'Step 1: Grab the slides collection for the slide.
With PPTSlide.Tags
    'Loop through each tag in the collection and print the name and value
    For i = 1 To .Count
        Debug.Print "Name = " & .Name(i)
        Debug.Print "Value = " & .Value(i)
    Next
End With

'Change the background color of the slide.
With PPTSlide
    .FollowMasterBackground = False
    .Background.Fill.BackColor.ObjectThemeColor = msoThemeColorBackground2
End With

'Copy a slide
PPTSlide.Copy

'Paste a slide in the presentation
ActivePresentation.Slides.Paste

'Duplicate a slide
PPTSlide.Duplicate

'More design stuff.
PPTSlide.BackgroundStyle = msoBackgroundStylePreset9

'Work with a hyperlink on the slide, first get the address.
Debug.Print PPTSlide.Hyperlinks.Item(1).Address

'Then follow the link
PPTSlide.Hyperlinks.Item(1).Follow

'Who is the parent of this slide?
Debug.Print PPTSlide.Parent.Name

End Sub

Sub AddHyperLink()

'Declare the variables.
Dim PPTShape As Shape
Dim PPTSlide As Slide

'First reference the slide.
Set PPTSlide = ActivePresentation.Slides(1)

'Let's add a hyperlink, this has to be done in two steps. First add a shape.
Set PPTShape = PPTSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=30, Top:=30, Width:=300, Height:=100)

    'Put some text in the text box, this will be overridden.
    PPTShape.TextFrame.TextRange.Text = "www.google.com"
    
    'Add an action setting, then specify it should be a hyperlink.
    With PPTShape.TextFrame.TextRange.ActionSettings(ppMouseClick).Hyperlink
    
        'Add the properties to the hyperlink.
        .Address = "https://www.google.com"
        .ScreenTip = "This is a link to google."
        .TextToDisplay = "My link to google."
        
    End With
End Sub
